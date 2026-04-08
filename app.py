import streamlit as st
import streamlit.components.v1 as components
import requests
import re
from datetime import date, timedelta
from calendar import monthrange

st.set_page_config(page_title="キャッシュフロー自動生成", page_icon="🏦", layout="wide")
st.title("🏦 freee キャッシュフロー自動生成")
st.caption("freee会計のデータから自動でキャッシュフロー表を作成します")

FREEE_BASE = "https://api.freee.co.jp/api/1"
TOKEN_URL  = "https://accounts.secure.freee.co.jp/public_api/token"

# ============================================================
# トークン自動更新
# ============================================================
def refresh_access_token():
    """
    Streamlit CloudのSecretsからリフレッシュトークンを使って
    アクセストークンを自動更新する
    """
    client_id     = st.secrets.get("FREEE_CLIENT_ID", "")
    client_secret = st.secrets.get("FREEE_CLIENT_SECRET", "")
    refresh_token = st.secrets.get("FREEE_REFRESH_TOKEN", "")

    if not all([client_id, client_secret, refresh_token]):
        return None

    try:
        res = requests.post(TOKEN_URL, data={
            "grant_type":    "refresh_token",
            "client_id":     client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
        }, timeout=10)
        res.raise_for_status()
        data = res.json()
        return data.get("access_token")
    except Exception as e:
        st.warning(f"トークン自動更新エラー: {e}")
        return None

# ============================================================
# freee API ヘルパー
# ============================================================
def freee_get(path, token, company_id, params=None):
    p = {"company_id": company_id}
    if params:
        p.update(params)
    res = requests.get(
        f"{FREEE_BASE}{path}",
        headers={"Authorization": f"Bearer {token}"},
        params=p, timeout=30,
    )
    res.raise_for_status()
    return res.json()

def get_companies(token):
    res = requests.get(f"{FREEE_BASE}/companies",
                       headers={"Authorization": f"Bearer {token}"}, timeout=10)
    res.raise_for_status()
    return res.json().get("companies", [])

def get_bank_walletables(token, company_id):
    d = freee_get("/walletables", token, company_id)
    return [w for w in d.get("walletables", []) if w.get("type") == "bank_account"]

def get_account_item_id_from_ledger(token, company_id, bank_names, start_date, end_date):
    """
    general_ledgersから銀行口座名に対応するaccount_item_idを取得
    """
    d = freee_get("/reports/general_ledgers", token, company_id, {
        "start_date": start_date,
        "end_date":   end_date,
    })
    rows = d.get("general_ledgers", [])
    result = {}
    for row in rows:
        if row["account_item_name"] in bank_names:
            result[row["account_item_name"]] = row["account_item_id"]
    return result

def get_deals_by_account(token, company_id, account_item_id, start_date, end_date):
    """
    deals APIをaccount_item_idで絞り込んで取得
    """
    all_deals = []
    offset = 0
    while True:
        d = freee_get("/deals", token, company_id, {
            "account_item_id":  account_item_id,
            "start_issue_date": start_date,
            "end_issue_date":   end_date,
            "offset":           offset,
            "limit":            100,
        })
        deals = d.get("deals", [])
        all_deals.extend(deals)
        if len(deals) < 100:
            break
        offset += 100
    return all_deals

def get_manual_journals_by_account(token, company_id, account_item_id, start_date, end_date):
    """
    manual_journals APIをaccount_item_idで絞り込んで取得
    """
    all_journals = []
    offset = 0
    while True:
        d = freee_get("/manual_journals", token, company_id, {
            "account_item_id":  account_item_id,
            "start_issue_date": start_date,
            "end_issue_date":   end_date,
            "offset":           offset,
            "limit":            100,
        })
        journals = d.get("manual_journals", [])
        all_journals.extend(journals)
        if len(journals) < 100:
            break
        offset += 100
    return all_journals

def get_walletable_balance(token, company_id, walletable_id, target_date):
    try:
        d = freee_get("/reports/walletable_balance", token, company_id, {
            "walletable_type": "bank_account",
            "walletable_id":   walletable_id,
            "date":            target_date,
        })
        return d.get("walletable_balance", {}).get("balance", 0) or 0
    except:
        return 0

# ============================================================
# 仕訳からCF行を抽出
# ============================================================
def extract_from_deal(deal, bank_account_item_ids):
    """
    取引（deal）から銀行が動いた行を抽出してCF行リストに変換
    """
    results = []
    issue_date   = deal.get("issue_date", "")
    partner_name = deal.get("partner_name") or ""
    details      = deal.get("details", [])

    bank_lines  = [d for d in details if d.get("account_item_id") in bank_account_item_ids]
    other_lines = [d for d in details if d.get("account_item_id") not in bank_account_item_ids]

    for bl in bank_lines:
        entry_side = bl.get("entry_side", "")
        amount     = bl.get("amount", 0) or 0
        net        = amount if entry_side == "debit" else -amount

        # 相手科目・取引先を特定
        counter_account = ""
        counter_partner = bl.get("partner_name") or partner_name
        if other_lines:
            # 金額が最大の相手行を採用
            best = max(other_lines, key=lambda x: x.get("amount", 0))
            counter_account = best.get("account_item_name") or ""
            if best.get("partner_name"):
                counter_partner = best.get("partner_name")

        results.append({
            "date":        issue_date,
            "amount":      net,
            "partner":     counter_partner,
            "account":     counter_account,
            "description": deal.get("description") or "",
            "source":      "deal",
        })
    return results

def extract_from_manual_journal(journal, bank_account_item_ids):
    """
    振替伝票（manual_journal）から銀行が動いた行を抽出してCF行リストに変換
    """
    results = []
    issue_date = journal.get("issue_date", "")
    details    = journal.get("details", [])

    bank_lines  = [d for d in details if d.get("account_item_id") in bank_account_item_ids]
    other_lines = [d for d in details if d.get("account_item_id") not in bank_account_item_ids]

    for bl in bank_lines:
        entry_side = bl.get("entry_side", "")
        amount     = bl.get("amount", 0) or 0
        net        = amount if entry_side == "debit" else -amount

        counter_account = ""
        counter_partner = bl.get("partner_name") or ""
        if other_lines:
            best = max(other_lines, key=lambda x: x.get("amount", 0))
            counter_account = best.get("account_item_name") or ""
            if best.get("partner_name") and not counter_partner:
                counter_partner = best.get("partner_name")

        results.append({
            "date":        issue_date,
            "amount":      net,
            "partner":     counter_partner,
            "account":     counter_account,
            "description": bl.get("description") or journal.get("description") or "",
            "source":      "manual_journal",
        })
    return results

# ============================================================
# 分類ロジック
# ============================================================
CORPORATE_KW = [
    "株式会社", "合同会社", "有限会社", "合名会社", "合資会社",
    "一般社団", "公益社団", "NPO", "社団法人", "財団法人",
    "事務所", "オフィス", "スタジオ", "ラボ", "工房", "商店",
    "inc", "Inc", "LLC", "ltd", "Ltd",
]

def is_personal_name(name):
    if not name:
        return False
    for kw in CORPORATE_KW:
        if kw in name:
            return False
    clean = name.replace("　", "").replace(" ", "")
    if re.fullmatch(r'[\u4e00-\u9fff]{2,4}', clean):
        return True
    if re.fullmatch(r'[\u3040-\u30ff\u4e00-\u9fff]{2,8}', clean):
        return True
    return False

# 相手科目名 → CF区分
ACCOUNT_MAP = {
    # 売上の入金
    "売掛金":         "売上の入金",
    "前受金":         "売上の入金",
    "未収入金":       "売上の入金",
    "未収利息":       "売上の入金",
    "受取利息":       "売上の入金",
    "雑収入":         "売上の入金",
    "預り金":         "売上の入金",
    # 原価（前方一致で判定）
    # → classify_row内でstartswith("仕入高")で処理
    # 広告宣伝費
    "広告宣伝費":     "広告宣伝費",
    # 人件費
    "役員報酬":       "人件費",
    "給与手当":       "人件費",
    "給与":           "人件費",
    "賞与":           "人件費",
    "法定福利費":     "人件費",
    "社会保険料":     "人件費",
    "労働保険料":     "人件費",
    # 税金
    "法人税等":       "税金",
    "租税公課":       "税金",
    "源泉所得税":     "税金",
    # 借入
    "長期借入金":     "_借入",
    "短期借入金":     "_借入",
    # 支払利息
    "支払利息":       "借入の返済",
    "未払利息":       "借入の返済",
    "長期未払金":     "借入の返済",
    # 貸付
    "長期貸付金":     "_貸付",
    "短期貸付金":     "_貸付",
    # 外注（個人名判定が必要）
    "外注費":         "_外注",
    "業務委託費":     "_外注",
    # 販管費
    "顧問料":         "販管費",
    "採用教育費":     "販管費",
    "研修費":         "販管費",
    "交際費":         "販管費",
    "会議費":         "販管費",
    "旅費交通費":     "販管費",
    "通信費":         "販管費",
    "消耗品費":       "販管費",
    "水道光熱費":     "販管費",
    "支払手数料":     "販管費",
    "システム利用料": "販管費",
    "地代家賃":       "販管費",
    "賃借料":         "販管費",
    "リース料":       "販管費",
    "保険料":         "販管費",
    "支払報酬料":     "販管費",
    "研究開発費":     "販管費",
    "新聞図書費":     "販管費",
    "諸会費":         "販管費",
    "荷造運賃":       "販管費",
    "車両費":         "販管費",
    "修繕費":         "販管費",
    "寄付金":         "販管費",
    "差入保証金":     "販管費",
    # 要遡及（未払金系）
    "未払金":         "_要遡及",
    "未払費用":       "_要遡及",
    "BD興行未払金":   "_要遡及",
    "スクール未払金": "_要遡及",
    "立替金":         "_要遡及",
    "仮払金":         "_要遡及",
    "経費精算":       "_要遡及",
    "前払費用":       "_要遡及",
    "仮受金":         "_要遡及",
}

def classify_row(row, skip_lookup=False):
    account = row.get("account", "")
    partner = row.get("partner", "")
    amount  = row.get("amount", 0)

    # 仕入高系（前方一致）
    if account.startswith("仕入高"):
        return "原価"
    # 売上高系が直接相手科目になる場合
    if account.startswith("売上高"):
        return "売上の入金"

    cat = ACCOUNT_MAP.get(account)

    if cat == "_借入":
        return "借入による収入" if amount > 0 else "借入の返済"
    if cat == "_貸付":
        return "貸付の回収" if amount > 0 else "貸付による支出"
    if cat == "_外注":
        return "人件費" if is_personal_name(partner) else "販管費"
    if cat == "_要遡及":
        return None if not skip_lookup else "販管費"
    if cat:
        return cat

    # マップにない科目
    return "売上の入金" if amount > 0 else "販管費"

# 取引先ごとの過去分類キャッシュ（セッション内で保持）
def get_past_category(token, company_id, bank_account_item_ids, partner, before_date):
    if not partner:
        return "販管費"

    cache = st.session_state.setdefault("partner_cache", {})
    if partner in cache:
        return cache[partner]

    end_dt   = date.fromisoformat(before_date) - timedelta(days=1)
    start_dt = end_dt - timedelta(days=90)

    try:
        for acct_id in list(bank_account_item_ids)[:2]:
            deals = get_deals_by_account(
                token, company_id, acct_id,
                start_dt.isoformat(), end_dt.isoformat()
            )
            for deal in deals:
                lines = extract_from_deal(deal, bank_account_item_ids)
                for line in lines:
                    if line.get("partner") == partner:
                        cat = classify_row(line, skip_lookup=True)
                        if cat:
                            cache[partner] = cat
                            return cat
    except:
        pass

    cache[partner] = "販管費"
    return "販管費"

def aggregate_cf(rows, token, company_id, bank_account_item_ids, start_date):
    cats = {
        "売上の入金": 0, "原価": 0, "広告宣伝費": 0, "販管費": 0,
        "人件費": 0, "税金": 0, "借入による収入": 0, "貸付の回収": 0,
        "貸付による支出": 0, "借入の返済": 0,
    }
    unclassified = []

    for row in rows:
        cat = classify_row(row)
        if cat is None:
            # 過去仕訳を遡って分類
            cat = get_past_category(
                token, company_id, bank_account_item_ids,
                row.get("partner", ""), start_date
            )
            row["_resolved"] = cat
            unclassified.append(row)
        cats[cat] += row.get("amount", 0)

    収入計    = cats["売上の入金"]
    支出計    = sum(cats[k] for k in ["原価", "広告宣伝費", "販管費", "人件費", "税金"])
    経常収支   = 収入計 + 支出計
    財務収入計  = cats["借入による収入"] + cats["貸付の回収"]
    財務支出計  = cats["貸付による支出"] + cats["借入の返済"]
    財務収支   = 財務収入計 + 財務支出計
    netCF    = 経常収支 + 財務収支

    return {
        **cats,
        "収入計": 収入計, "支出計": 支出計, "経常収支": 経常収支,
        "財務収入計": 財務収入計, "財務支出計": 財務支出計,
        "財務収支": 財務収支, "netCF": netCF,
        "_unclassified": unclassified,
        "_all_rows": rows,
    }

# ============================================================
# HTML生成
# ============================================================
def generate_html(cf_data, company_name, months, bank_names, verify_data):
    today    = date.today().isoformat()
    period   = f"{months[0]['year']}年{months[0]['month']}月 〜 {months[-1]['year']}年{months[-1]['month']}月"
    acct_str = "・".join(bank_names)
    n        = len(months)

    def fmt(v):
        if not v:
            return '<span class="v-dash">—</span>'
        if v < 0:
            return f'<span class="v-exp">▲ {abs(int(v)):,}</span>'
        return f'<span class="v-inc">{int(v):,}</span>'

    def fmt_num(v):
        return f"{int(v):,}" if v else "—"

    def fmt_diff(v):
        return '<span class="v-ok">0 ✓</span>' if v == 0 else f'<span class="v-ng">▲ {abs(int(v)):,} ✗</span>'

    def mk(m):
        return str(m["year"]) + "-" + str(m["month"])

    def cells(key):
        return "".join(
            '<td class="num">' + fmt(cf_data.get(mk(m), {}).get(key, 0)) + '</td>'
            for m in months)

    def bal_cells(key):
        return "".join(
            '<td class="num bg-bal">' + fmt_num(cf_data.get(mk(m), {}).get(key, 0)) + '</td>'
            for m in months)

    def total_cells(key):
        parts = []
        for m in months:
            v   = cf_data.get(mk(m), {}).get(key, 0)
            cls = "v-neg" if v < 0 else "v-pos" if v > 0 else "v-zero"
            parts.append(f'<td class="num bg-gry {cls}">' + (fmt(v) if v != 0 else "0") + '</td>')
        return "".join(parts)

    cards = ""
    for m in months:
        d       = cf_data.get(mk(m), {})
        net     = d.get("netCF", 0)
        closing = d.get("closingBalance", 0) or 0
        tag = ('<span class="tag-zero">± 0</span>' if net == 0
               else f'<span class="tag-pos">▲ {abs(int(net)):,}</span>' if net > 0
               else f'<span class="tag-neg">▼ {abs(int(net)):,}</span>')
        cards += f"""
    <div class="card">
      <div class="card-month">{m['year']}年{m['month']}月</div>
      <div class="card-balance"><span class="yen">¥</span>{int(closing):,}</div>
      <div class="card-sep"></div>
      <div class="card-footer"><span class="card-footer-label">月次収支</span>{tag}</div>
    </div>"""

    vf = "".join(
        '<td class="num bg-vrf">' + fmt_num(verify_data.get(mk(m), 0)) + '</td>'
        for m in months)
    vd = "".join(
        '<td class="num bg-vrf">' + fmt_diff(
            (cf_data.get(mk(m), {}).get("closingBalance", 0) or 0) -
            (verify_data.get(mk(m), 0) or 0)
        ) + '</td>'
        for m in months)
    col_headers = "".join(f"<th>{m['year']}年{m['month']}月</th>" for m in months)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>{company_name}｜キャッシュフロー表 {period}</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Noto Sans JP','Meiryo',sans-serif;background:#f0f2f5;color:#1a1a1a;font-size:13px}}
  .page-header{{background:#fff;border-bottom:1px solid #d0d5dd;padding:20px 32px;display:flex;align-items:flex-end;justify-content:space-between}}
  .page-header .company{{font-size:18px;font-weight:700}}
  .page-header .subtitle{{font-size:12px;color:#667085;margin-top:3px}}
  .page-header .meta{{font-size:11px;color:#98a2b3;text-align:right}}
  .container{{max-width:980px;margin:24px auto;padding:0 20px 60px}}
  .cards{{display:grid;grid-template-columns:repeat({n},1fr);gap:12px;margin-bottom:24px}}
  .card{{background:#fff;border:1px solid #d0d5dd;border-radius:6px;padding:16px 18px;border-top:3px solid #4a90d9}}
  .card .card-month{{font-size:12px;font-weight:700;color:#344054;margin-bottom:6px}}
  .card .card-balance{{font-size:20px;font-weight:700;margin-bottom:8px}}
  .card .card-balance .yen{{font-size:13px;color:#667085;margin-right:2px;font-weight:400}}
  .card .card-sep{{height:1px;background:#e4e7ec;margin:8px 0}}
  .card .card-footer{{display:flex;justify-content:space-between;font-size:11px}}
  .card .card-footer-label{{color:#98a2b3}}
  .tag-neg{{background:#fff1f1;border:1px solid #fca5a5;color:#b91c1c;padding:1px 7px;border-radius:3px;font-size:11px;font-weight:700}}
  .tag-pos{{background:#f0fdf4;border:1px solid #86efac;color:#15803d;padding:1px 7px;border-radius:3px;font-size:11px;font-weight:700}}
  .tag-zero{{background:#f9fafb;border:1px solid #d0d5dd;color:#98a2b3;padding:1px 7px;border-radius:3px;font-size:11px;font-weight:700}}
  .table-wrap{{background:#fff;border:1px solid #d0d5dd;border-radius:6px;overflow:hidden}}
  table{{width:100%;border-collapse:collapse}}
  thead th{{background:#f2f4f7;border:1px solid #d0d5dd;padding:9px 14px;font-size:12px;font-weight:700;color:#344054;text-align:center;white-space:nowrap}}
  thead th.col-label{{text-align:left;width:30%}}
  td.g-label{{writing-mode:vertical-rl;text-align:center;font-size:12px;font-weight:700;letter-spacing:.1em;white-space:nowrap;padding:0;border:1px solid #d0d5dd;width:24px}}
  td.g-keijo{{background:#cce0f5;color:#1a4a7a}}
  td.g-zaim{{background:#ffd6d6;color:#7a1a1a}}
  .r-normal td{{background:#fff;border:1px solid #e4e7ec;padding:8px 14px;font-size:13px;color:#344054}}
  .r-normal td.num{{text-align:right}}
  .r-sub-inc td{{background:#e8f3ff;border:1px solid #d0d5dd;padding:8px 14px;font-weight:700;font-size:13px}}
  .r-sub-inc td.num{{text-align:right}}
  .r-sub-exp td{{background:#fff0f0;border:1px solid #d0d5dd;padding:8px 14px;font-weight:700;font-size:13px}}
  .r-sub-exp td.num{{text-align:right}}
  .r-total td{{background:#e9ecef;border:1px solid #d0d5dd;padding:9px 14px;font-weight:700;font-size:13px}}
  .r-total td.num{{text-align:right}}
  .r-balance td{{background:#fff9db;border:1px solid #d0d5dd;padding:9px 14px;font-weight:700;font-size:13px}}
  .r-balance td.num{{text-align:right}}
  .r-verify td{{background:#f9fafb;border:1px solid #e4e7ec;padding:7px 14px;font-size:12px;color:#667085}}
  .r-verify td.num{{text-align:right}}
  .r-diff td{{background:#f9fafb;border:1px solid #e4e7ec;padding:7px 14px;font-size:12px;font-weight:700}}
  .r-diff td.num{{text-align:right}}
  .bg-inc{{background:#e8f3ff!important;border:1px solid #d0d5dd!important}}
  .bg-exp{{background:#fff0f0!important;border:1px solid #d0d5dd!important}}
  .bg-bal{{background:#fff9db!important;border:1px solid #d0d5dd!important}}
  .bg-gry{{background:#e9ecef!important;border:1px solid #d0d5dd!important}}
  .bg-vrf{{background:#f9fafb!important;border:1px solid #e4e7ec!important}}
  .v-inc{{color:#1558b0!important;font-weight:700}}
  .v-exp{{color:#c0392b!important}}
  .v-neg{{color:#c0392b!important;font-weight:700}}
  .v-pos{{color:#1b7f4a!important;font-weight:700}}
  .v-ok{{color:#1b7f4a!important;font-weight:700}}
  .v-ng{{color:#c0392b!important;font-weight:700}}
  .v-dash{{color:#bdc3ce!important}}
  .v-zero{{color:#98a2b3!important}}
  .footnote{{margin-top:16px;font-size:11px;color:#98a2b3;line-height:2;padding-left:12px;border-left:3px solid #e4e7ec}}
</style>
</head>
<body>
<div class="page-header">
  <div>
    <div class="company">{company_name}　キャッシュフロー表</div>
    <div class="subtitle">対象期間: {period}　｜　{acct_str}（預金勘定）　｜　単位: 円</div>
  </div>
  <div class="meta">出力日: {today}</div>
</div>
<div class="container">
  <div class="cards">{cards}</div>
  <div class="table-wrap">
  <table>
    <thead>
      <tr><th colspan="2"></th><th class="col-label">項目</th>{col_headers}</tr>
    </thead>
    <tbody>
      <tr class="r-balance">
        <td colspan="2" class="bg-bal"></td>
        <td class="bg-bal">月初繰越残高　(A)</td>{bal_cells("openingBalance")}
      </tr>
      <tr class="r-normal">
        <td class="g-label g-keijo" rowspan="8">経常収支</td>
        <td class="bg-inc"></td><td>売上の入金</td>{cells("売上の入金")}
      </tr>
      <tr class="r-sub-inc"><td class="bg-inc"></td><td>収入計　(B)</td>{cells("収入計")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>原価</td>{cells("原価")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>広告宣伝費</td>{cells("広告宣伝費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>販管費（外注費含む）</td>{cells("販管費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>人件費</td>{cells("人件費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>税金</td>{cells("税金")}</tr>
      <tr class="r-sub-exp"><td class="bg-exp"></td><td>支出計　(C)</td>{cells("支出計")}</tr>
      <tr class="r-total">
        <td colspan="2" class="bg-gry"></td>
        <td>経常収支　(D)=(B)-(C)</td>{total_cells("経常収支")}
      </tr>
      <tr class="r-normal">
        <td class="g-label g-zaim" rowspan="7">財務収支</td>
        <td class="bg-inc"></td><td>借入による収入</td>{cells("借入による収入")}
      </tr>
      <tr class="r-normal"><td class="bg-inc"></td><td>貸付の回収</td>{cells("貸付の回収")}</tr>
      <tr class="r-sub-inc"><td class="bg-inc"></td><td>財務収入計　(E)</td>{cells("財務収入計")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>貸付による支出</td>{cells("貸付による支出")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>借入の返済</td>{cells("借入の返済")}</tr>
      <tr class="r-sub-exp"><td class="bg-exp"></td><td>財務支出計　(F)</td>{cells("財務支出計")}</tr>
      <tr class="r-total">
        <td colspan="2" class="bg-gry"></td>
        <td>財務収支　(G)=(E)-(F)</td>{total_cells("財務収支")}
      </tr>
      <tr class="r-total">
        <td colspan="2" class="bg-gry"></td>
        <td class="bg-gry">合計収支　(H)=(D)+(G)</td>{total_cells("netCF")}
      </tr>
      <tr class="r-balance">
        <td colspan="2" class="bg-bal"></td>
        <td class="bg-bal">次月繰越残高　(I)=(A)+(H)</td>{bal_cells("closingBalance")}
      </tr>
      <tr class="r-verify">
        <td colspan="2" class="bg-vrf"></td>
        <td class="bg-vrf">freee実残高（{acct_str}）</td>{vf}
      </tr>
      <tr class="r-diff">
        <td colspan="2" class="bg-vrf"></td>
        <td class="bg-vrf" style="color:#98a2b3;">差異（0=一致）</td>{vd}
      </tr>
    </tbody>
  </table>
  </div>
  <div class="footnote">
    ※ 取引（deals）・振替伝票（manual_journals）から銀行口座が動いた仕訳を抽出して集計しています。<br>
    ※ 外注費・業務委託費は取引先名で個人名判定し、個人名の場合は人件費に分類しています。<br>
    ※ 未払金等で分類不明の場合は同取引先の過去3ヶ月仕訳を参照して分類しています。<br>
    ※ 出力日: {today}
  </div>
</div>
</body>
</html>"""

# ============================================================
# Streamlit UI
# ============================================================
with st.sidebar:
    st.header("⚙️ 設定")

    # リフレッシュトークンで自動更新を試みる
    saved_token = st.secrets.get("FREEE_TOKEN", "")
    auto_token  = None

    if st.secrets.get("FREEE_REFRESH_TOKEN"):
        auto_token = refresh_access_token()

    if auto_token:
        st.success("✅ トークン自動更新済み")
        token = auto_token
    elif saved_token:
        st.success("✅ トークン読み込み済み")
        use_saved = st.checkbox("保存済みトークンを使用", value=True)
        token = saved_token if use_saved else st.text_input("freee アクセストークン", type="password")
    else:
        token = st.text_input("freee アクセストークン", type="password", placeholder="トークンを入力")

if not token:
    st.info("👈 左のサイドバーにfreeeのアクセストークンを入力してください")
    st.stop()

@st.cache_data(ttl=300)
def load_companies(t):
    return get_companies(t)

try:
    companies = load_companies(token)
except Exception as e:
    st.error(f"接続エラー: {e}")
    st.stop()

with st.sidebar:
    company_options = {c["display_name"]: c["id"] for c in companies}
    selected_name   = st.selectbox("事業所", list(company_options.keys()))
    company_id      = company_options[selected_name]
    st.divider()
    st.subheader("📅 対象期間")
    c1, c2 = st.columns(2)
    with c1:
        years       = list(range(2020, date.today().year + 1))
        start_year  = st.selectbox("開始年", years, index=len(years)-1)
        start_month = st.selectbox("開始月", range(1, 13), index=0)
    with c2:
        end_year  = st.selectbox("終了年", years, index=len(years)-1)
        end_month = st.selectbox("終了月", range(1, 13), index=2)

    generate_btn = st.button("🚀 キャッシュフロー生成", use_container_width=True, type="primary")
    st.divider()
    debug_mode = st.checkbox("🔍 デバッグモード", value=False)

if "html_result" not in st.session_state:
    st.session_state["html_result"] = None

if generate_btn:
    st.session_state["partner_cache"] = {}  # キャッシュリセット
    months = []
    cy, cm = start_year, start_month
    while (cy, cm) <= (end_year, end_month):
        months.append({"year": cy, "month": cm})
        cm += 1
        if cm > 12:
            cm = 1; cy += 1

    progress = st.progress(0, text="銀行口座を取得中...")

    # 銀行口座取得
    try:
        bank_accounts = get_bank_walletables(token, company_id)
    except Exception as e:
        st.error(f"口座取得エラー: {e}")
        st.stop()

    if not bank_accounts:
        st.error("銀行口座が見つかりません")
        st.stop()

    bank_names = [b["name"] for b in bank_accounts]

    # general_ledgersからaccount_item_idを取得
    progress.progress(5, text="勘定科目IDを取得中...")
    first_mon  = months[0]
    last_mon   = months[-1]
    last_day_f = monthrange(first_mon["year"], first_mon["month"])[1]
    s_date_f   = f"{first_mon['year']}-{first_mon['month']:02d}-01"
    e_date_f   = f"{last_mon['year']}-{last_mon['month']:02d}-{monthrange(last_mon['year'], last_mon['month'])[1]}"

    try:
        bank_id_map = get_account_item_id_from_ledger(
            token, company_id, bank_names, s_date_f, e_date_f
        )
    except Exception as e:
        st.error(f"勘定科目ID取得エラー: {e}")
        st.stop()

    if not bank_id_map:
        st.error("銀行口座の勘定科目IDが取得できませんでした")
        st.stop()

    bank_account_item_ids = set(bank_id_map.values())

    if debug_mode:
        st.info(f"銀行口座IDマップ: {bank_id_map}")

    cf_data     = {}
    verify_data = {}

    for i, mon in enumerate(months):
        key      = str(mon["year"]) + "-" + str(mon["month"])
        last_day = monthrange(mon["year"], mon["month"])[1]
        s_date   = f"{mon['year']}-{mon['month']:02d}-01"
        e_date   = f"{mon['year']}-{mon['month']:02d}-{last_day}"
        pct      = int(10 + (i / len(months)) * 70)

        all_rows = []

        for bank_name, acct_id in bank_id_map.items():
            progress.progress(pct, text=f"{mon['year']}年{mon['month']}月 / {bank_name} 取引取得中...")
            try:
                deals = get_deals_by_account(token, company_id, acct_id, s_date, e_date)
                for deal in deals:
                    all_rows.extend(extract_from_deal(deal, bank_account_item_ids))
            except Exception as e:
                st.warning(f"deals {bank_name}: {e}")

            try:
                journals = get_manual_journals_by_account(token, company_id, acct_id, s_date, e_date)
                for journal in journals:
                    all_rows.extend(extract_from_manual_journal(journal, bank_account_item_ids))
            except Exception as e:
                st.warning(f"manual_journals {bank_name}: {e}")

        progress.progress(pct + 5, text=f"{mon['year']}年{mon['month']}月 分類中...")
        agg = aggregate_cf(all_rows, token, company_id, bank_account_item_ids, s_date)

        # 月末残高
        closing = sum(
            get_walletable_balance(token, company_id, b["id"], e_date)
            for b in bank_accounts
        )
        agg["closingBalance"] = closing
        agg["openingBalance"] = 0
        cf_data[key]     = agg
        verify_data[key] = closing

    # 月初残高を連鎖
    for i in range(1, len(months)):
        pk = str(months[i-1]["year"]) + "-" + str(months[i-1]["month"])
        ck = str(months[i]["year"])   + "-" + str(months[i]["month"])
        if pk in cf_data and ck in cf_data:
            cf_data[ck]["openingBalance"] = cf_data[pk]["closingBalance"]

    progress.progress(95, text="HTML生成中...")
    html = generate_html(cf_data, selected_name, months, bank_names, verify_data)

    st.session_state.update({
        "html_result": html,
        "cf_data":     cf_data,
        "months":      months,
        "verify_data": verify_data,
        "period_str":  f"{start_year}{start_month:02d}-{end_year}{end_month:02d}",
    })
    progress.progress(100, text="完了！")

# ============================================================
# 結果表示
# ============================================================
if st.session_state.get("html_result"):
    html        = st.session_state["html_result"]
    cf_data     = st.session_state["cf_data"]
    months      = st.session_state["months"]
    verify_data = st.session_state.get("verify_data", {})
    period_str  = st.session_state.get("period_str", "output")

    st.success("✅ キャッシュフロー表が生成されました！内容を確認してからダウンロードしてください。")

    has_diff = False
    for mon in months:
        key = str(mon["year"]) + "-" + str(mon["month"])
        cb  = cf_data.get(key, {}).get("closingBalance", 0) or 0
        fb  = verify_data.get(key, 0) or 0
        if abs(cb - fb) > 0 and fb > 0:
            st.warning(f"⚠ {mon['year']}年{mon['month']}月: freee残高との差異 {cb - fb:,}円")
            has_diff = True
    if not has_diff:
        st.info("✅ 全月、freee残高との差異なし")

    st.subheader("📊 プレビュー")
    components.html(html, height=800, scrolling=True)
    st.divider()

    st.download_button(
        label="⬇️ HTMLをダウンロード",
        data=html.encode("utf-8"),
        file_name=f"cashflow_{period_str}.html",
        mime="text/html",
        use_container_width=True,
        type="primary",
        on_click=lambda: None,  # ページリセット防止
    )

    if debug_mode:
        for mon in months:
            key = str(mon["year"]) + "-" + str(mon["month"])
            d   = cf_data.get(key, {})
            rows = d.get("_all_rows", [])
            if rows:
                with st.expander(f"📋 {mon['year']}年{mon['month']}月 仕訳明細（{len(rows)}件）"):
                    import pandas as pd
                    st.dataframe(pd.DataFrame(rows), use_container_width=True)
            uncl = d.get("_unclassified", [])
            if uncl:
                with st.expander(f"⚠ {mon['year']}年{mon['month']}月 遡及処理した仕訳（{len(uncl)}件）"):
                    import pandas as pd
                    st.dataframe(pd.DataFrame(uncl), use_container_width=True)
