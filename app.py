import streamlit as st
import requests
import json
from datetime import date
from calendar import monthrange

st.set_page_config(page_title="キャッシュフロー自動生成", page_icon="🏦", layout="wide")
st.title("🏦 freee キャッシュフロー自動生成")
st.caption("freee会計のデータから自動でキャッシュフロー表を作成します")

FREEE_BASE = "https://api.freee.co.jp/api/1"

# ============================================================
# freee API ヘルパー
# ============================================================
def freee_get(path, token, company_id, params=None):
    p = params or {}
    p["company_id"] = company_id
    res = requests.get(
        f"{FREEE_BASE}{path}",
        headers={"Authorization": f"Bearer {token}"},
        params=p
    )
    res.raise_for_status()
    return res.json()

def get_companies(token):
    res = requests.get(
        "https://api.freee.co.jp/api/1/companies",
        headers={"Authorization": f"Bearer {token}"}
    )
    res.raise_for_status()
    return res.json().get("companies", [])

def get_deposit_accounts(token, company_id):
    # 口座一覧APIから「銀行口座」カテゴリのみ取得
    d = freee_get("/walletables", token, company_id, {"type": "bank"})
    walletables = d.get("walletables", [])

    # デバッグ用に保存
    st.session_state["all_account_items"] = [
        {"id": w["id"], "name": w["name"], "type": w.get("type", ""), "account_item_id": w.get("account_item_id", "")}
        for w in walletables
    ]

    # 勘定科目IDと名前のペアに変換（general_ledger APIで使うaccount_item_idが必要）
    result = []
    for w in walletables:
        if w.get("account_item_id"):
            result.append({
                "id":   w["account_item_id"],
                "name": w["name"],
            })
    return result

def get_general_ledger(token, company_id, account_item_id, start_date, end_date):
    all_rows = []
    offset = 0
    max_pages = 20  # 無限ループ防止
    while offset < max_pages * 100:
        d = freee_get("/reports/general_ledgers", token, company_id, {
            "account_item_id": account_item_id,
            "start_date": start_date,
            "end_date": end_date,
            "offset": offset,
            "limit": 100,
        })
        # レスポンス構造をデバッグ用に保存
        if offset == 0:
            st.session_state["last_ledger_response"] = d

        # freee APIのレスポンス構造に対応（複数パターン）
        rows = (
            d.get("account_item", {}).get("balances") or
            d.get("balances") or
            d.get("general_ledgers") or
            []
        )
        all_rows.extend(rows)
        if len(rows) < 100:
            break
        offset += 100
    return all_rows

# ============================================================
# 分類ロジック
# ============================================================
def is_personal_name(name):
    if not name:
        return False
    corporate_kw = ["株式会社", "合同会社", "有限会社", "合名会社", "合資会社",
                    "一般社団", "公益社団", "NPO", "社団法人", "財団法人",
                    "事務所", "オフィス", "スタジオ", "ラボ", "工房", "商店", "inc", "Inc", "LLC"]
    for kw in corporate_kw:
        if kw in name:
            return False
    import re
    clean = name.replace("　", "").replace(" ", "")
    if re.fullmatch(r'[\u4e00-\u9fff]{2,4}', clean):
        return True
    return False

def classify_journal(j):
    account = j.get("account", "")
    partner = j.get("partner", "")
    amount  = j.get("amount", 0)

    if amount > 0:
        for kw in ["売上", "売掛", "受取利息", "雑収入", "前受"]:
            if kw in account:
                return "売上の入金"
        for kw in ["短期借入", "長期借入"]:
            if kw in account:
                return "借入による収入"
        for kw in ["貸付金"]:
            if kw in account and "回収" in account:
                return "貸付の回収"
        return "売上の入金"

    for kw in ["仕入", "原価", "材料"]:
        if kw in account:
            return "原価"
    for kw in ["広告宣伝"]:
        if kw in account:
            return "広告宣伝費"
    for kw in ["給与", "賞与", "役員報酬", "社会保険", "労働保険"]:
        if kw in account:
            return "人件費"
    for kw in ["法人税", "消費税", "源泉", "住民税", "租税"]:
        if kw in account:
            return "税金"
    for kw in ["支払利息", "借入金返済", "長期借入金", "短期借入金"]:
        if kw in account:
            return "借入の返済"
    for kw in ["貸付"]:
        if kw in account and "回収" not in account:
            return "貸付による支出"
    if "外注" in account or "業務委託" in account:
        return "人件費" if is_personal_name(partner) else "販管費"
    if "未払" in account or "立替" in account:
        return "人件費" if is_personal_name(partner) else "販管費"
    return "販管費"

def aggregate(journals):
    cats = {
        "売上の入金": 0, "原価": 0, "広告宣伝費": 0, "販管費": 0,
        "人件費": 0, "税金": 0, "借入による収入": 0, "貸付の回収": 0,
        "貸付による支出": 0, "借入の返済": 0,
    }
    for j in journals:
        cat = classify_journal(j)
        cats[cat] += j.get("amount", 0)

    収入計    = cats["売上の入金"]
    支出計    = sum(cats[k] for k in ["原価", "広告宣伝費", "販管費", "人件費", "税金"])
    経常収支   = 収入計 + 支出計
    財務収入計  = cats["借入による収入"] + cats["貸付の回収"]
    財務支出計  = cats["貸付による支出"] + cats["借入の返済"]
    財務収支   = 財務収入計 + 財務支出計
    netCF    = 経常収支 + 財務収支

    return {**cats,
            "収入計": 収入計, "支出計": 支出計, "経常収支": 経常収支,
            "財務収入計": 財務収入計, "財務支出計": 財務支出計,
            "財務収支": 財務収支, "netCF": netCF}

# ============================================================
# HTML生成
# ============================================================
def generate_html(cf_data, company_name, months, account_names, verify_data):
    today     = date.today().isoformat()
    period    = f"{months[0]['year']}年{months[0]['month']}月 〜 {months[-1]['year']}年{months[-1]['month']}月"
    acct_str  = "・".join(account_names)
    n         = len(months)

    def fmt(v):
        if v is None or v == 0:
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
        return "".join('<td class="num">' + fmt(cf_data.get(mk(m), {}).get(key, 0)) + '</td>' for m in months)

    def bal_cells(key):
        return "".join('<td class="num bg-bal">' + fmt_num(cf_data.get(mk(m), {}).get(key, 0)) + '</td>' for m in months)

    def total_cells(key):
        parts = []
        for m in months:
            v   = cf_data.get(mk(m), {}).get(key, 0)
            cls = "v-neg" if v < 0 else "v-pos" if v > 0 else "v-zero"
            parts.append(f'<td class="num bg-gry {cls}">' + (fmt(v) if v != 0 else "0") + '</td>')
        return "".join(parts)

    cards = ""
    for m in months:
        d      = cf_data.get(mk(m), {})
        net    = d.get("netCF", 0)
        closing = d.get("closingBalance", 0)
        tag    = ('<span class="tag-zero">± 0</span>' if net == 0
                  else f'<span class="tag-pos">▲ {abs(int(net)):,}</span>' if net > 0
                  else f'<span class="tag-neg">▼ {abs(int(net)):,}</span>')
        cards += f"""
    <div class="card">
      <div class="card-month">{m['year']}年{m['month']}月</div>
      <div class="card-balance"><span class="yen">¥</span>{int(closing):,}</div>
      <div class="card-sep"></div>
      <div class="card-footer"><span class="card-footer-label">月次収支</span>{tag}</div>
    </div>"""

    vf = "".join('<td class="num bg-vrf">' + fmt_num(verify_data.get(mk(m), 0)) + '</td>' for m in months)
    vd_parts = []
    for m in months:
        cb = cf_data.get(mk(m), {}).get("closingBalance", 0)
        fb = verify_data.get(mk(m), 0)
        vd_parts.append('<td class="num bg-vrf">' + fmt_diff(cb - fb) + '</td>')
    vd = "".join(vd_parts)

    col_headers = "".join(f"<th>{m['year']}年{m['month']}月</th>" for m in months)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{company_name}｜キャッシュフロー表 {period}</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Noto Sans JP', 'Meiryo', sans-serif; background: #f0f2f5; color: #1a1a1a; font-size: 13px; }}
  .page-header {{ background: #fff; border-bottom: 1px solid #d0d5dd; padding: 20px 32px; display: flex; align-items: flex-end; justify-content: space-between; }}
  .page-header .company {{ font-size: 18px; font-weight: 700; }}
  .page-header .subtitle {{ font-size: 12px; color: #667085; margin-top: 3px; }}
  .page-header .meta {{ font-size: 11px; color: #98a2b3; text-align: right; }}
  .container {{ max-width: 980px; margin: 24px auto; padding: 0 20px 60px; }}
  .cards {{ display: grid; grid-template-columns: repeat({n}, 1fr); gap: 12px; margin-bottom: 24px; }}
  .card {{ background: #fff; border: 1px solid #d0d5dd; border-radius: 6px; padding: 16px 18px; border-top: 3px solid #4a90d9; }}
  .card .card-month {{ font-size: 12px; font-weight: 700; color: #344054; margin-bottom: 6px; }}
  .card .card-balance {{ font-size: 20px; font-weight: 700; margin-bottom: 8px; }}
  .card .card-balance .yen {{ font-size: 13px; color: #667085; margin-right: 2px; font-weight: 400; }}
  .card .card-sep {{ height: 1px; background: #e4e7ec; margin: 8px 0; }}
  .card .card-footer {{ display: flex; justify-content: space-between; font-size: 11px; }}
  .card .card-footer-label {{ color: #98a2b3; }}
  .tag-neg {{ background: #fff1f1; border: 1px solid #fca5a5; color: #b91c1c; padding: 1px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; }}
  .tag-pos {{ background: #f0fdf4; border: 1px solid #86efac; color: #15803d; padding: 1px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; }}
  .tag-zero {{ background: #f9fafb; border: 1px solid #d0d5dd; color: #98a2b3; padding: 1px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; }}
  .table-wrap {{ background: #fff; border: 1px solid #d0d5dd; border-radius: 6px; overflow: hidden; }}
  table {{ width: 100%; border-collapse: collapse; }}
  thead th {{ background: #f2f4f7; border: 1px solid #d0d5dd; padding: 9px 14px; font-size: 12px; font-weight: 700; color: #344054; text-align: center; white-space: nowrap; }}
  thead th.col-label {{ text-align: left; width: 30%; }}
  td.g-label {{ writing-mode: vertical-rl; text-align: center; font-size: 12px; font-weight: 700; letter-spacing: 0.1em; white-space: nowrap; padding: 0; border: 1px solid #d0d5dd; width: 24px; }}
  td.g-keijo {{ background: #cce0f5; color: #1a4a7a; }}
  td.g-zaim  {{ background: #ffd6d6; color: #7a1a1a; }}
  .r-normal td {{ background: #fff; border: 1px solid #e4e7ec; padding: 8px 14px; font-size: 13px; color: #344054; }}
  .r-normal td.num {{ text-align: right; }}
  .r-sub-inc td {{ background: #e8f3ff; border: 1px solid #d0d5dd; padding: 8px 14px; font-weight: 700; font-size: 13px; }}
  .r-sub-inc td.num {{ text-align: right; }}
  .r-sub-exp td {{ background: #fff0f0; border: 1px solid #d0d5dd; padding: 8px 14px; font-weight: 700; font-size: 13px; }}
  .r-sub-exp td.num {{ text-align: right; }}
  .r-total td {{ background: #e9ecef; border: 1px solid #d0d5dd; padding: 9px 14px; font-weight: 700; font-size: 13px; }}
  .r-total td.num {{ text-align: right; }}
  .r-balance td {{ background: #fff9db; border: 1px solid #d0d5dd; padding: 9px 14px; font-weight: 700; font-size: 13px; }}
  .r-balance td.num {{ text-align: right; }}
  .r-verify td {{ background: #f9fafb; border: 1px solid #e4e7ec; padding: 7px 14px; font-size: 12px; color: #667085; }}
  .r-verify td.num {{ text-align: right; }}
  .r-diff td {{ background: #f9fafb; border: 1px solid #e4e7ec; padding: 7px 14px; font-size: 12px; font-weight: 700; }}
  .r-diff td.num {{ text-align: right; }}
  .bg-inc {{ background: #e8f3ff !important; border: 1px solid #d0d5dd !important; }}
  .bg-exp {{ background: #fff0f0 !important; border: 1px solid #d0d5dd !important; }}
  .bg-bal {{ background: #fff9db !important; border: 1px solid #d0d5dd !important; }}
  .bg-gry {{ background: #e9ecef !important; border: 1px solid #d0d5dd !important; }}
  .bg-vrf {{ background: #f9fafb !important; border: 1px solid #e4e7ec !important; }}
  .v-inc  {{ color: #1558b0 !important; font-weight: 700; }}
  .v-exp  {{ color: #c0392b !important; }}
  .v-neg  {{ color: #c0392b !important; font-weight: 700; }}
  .v-pos  {{ color: #1b7f4a !important; font-weight: 700; }}
  .v-ok   {{ color: #1b7f4a !important; font-weight: 700; }}
  .v-ng   {{ color: #c0392b !important; font-weight: 700; }}
  .v-dash {{ color: #bdc3ce !important; }}
  .v-zero {{ color: #98a2b3 !important; }}
  .footnote {{ margin-top: 16px; font-size: 11px; color: #98a2b3; line-height: 2; padding-left: 12px; border-left: 3px solid #e4e7ec; }}
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
      <tr class="r-sub-inc">
        <td class="bg-inc"></td><td>収入計　(B)</td>{cells("収入計")}
      </tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>原価</td>{cells("原価")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>広告宣伝費</td>{cells("広告宣伝費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>販管費（外注費含む）</td>{cells("販管費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>人件費</td>{cells("人件費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>税金</td>{cells("税金")}</tr>
      <tr class="r-sub-exp">
        <td class="bg-exp"></td><td>支出計　(C)</td>{cells("支出計")}
      </tr>
      <tr class="r-total">
        <td colspan="2" class="bg-gry"></td>
        <td>経常収支　(D)=(B)-(C)</td>{total_cells("経常収支")}
      </tr>
      <tr class="r-normal">
        <td class="g-label g-zaim" rowspan="7">財務収支</td>
        <td class="bg-inc"></td><td>借入による収入</td>{cells("借入による収入")}
      </tr>
      <tr class="r-normal"><td class="bg-inc"></td><td>貸付の回収</td>{cells("貸付の回収")}</tr>
      <tr class="r-sub-inc">
        <td class="bg-inc"></td><td>財務収入計　(E)</td>{cells("財務収入計")}
      </tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>貸付による支出</td>{cells("貸付による支出")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>借入の返済</td>{cells("借入の返済")}</tr>
      <tr class="r-sub-exp">
        <td class="bg-exp"></td><td>財務支出計　(F)</td>{cells("財務支出計")}
      </tr>
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
    ※ 預金勘定の総勘定元帳をもとに集計しています。<br>
    ※ 人件費は社会保険料および個人名への業務委託費を含みます。<br>
    ※ 販管費（外注費含む）は法人・屋号への外注費・業務委託費を含みます。<br>
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

    # Secretsに保存済みのトークンを自動読み込み
    saved_token = st.secrets.get("FREEE_TOKEN", "")

    if saved_token:
        st.success("✅ トークン読み込み済み")
        use_saved = st.checkbox("保存済みトークンを使用", value=True)
        if use_saved:
            token = saved_token
        else:
            token = st.text_input("freee アクセストークン", type="password", placeholder="別のトークンを入力")
    else:
        token = st.text_input("freee アクセストークン", type="password", placeholder="トークンを入力")
        st.caption("freee開発者ページで発行したトークンを入力してください")

if not token:
    st.info("👈 左のサイドバーにfreeeのアクセストークンを入力してください")
    st.stop()

@st.cache_data(ttl=300)
def load_companies(token):
    return get_companies(token)

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
    col1, col2 = st.columns(2)
    with col1:
        years       = list(range(2020, date.today().year + 1))
        start_year  = st.selectbox("開始年", years, index=len(years) - 1)
        start_month = st.selectbox("開始月", range(1, 13), index=0)
    with col2:
        end_year    = st.selectbox("終了年", years, index=len(years) - 1)
        end_month   = st.selectbox("終了月", range(1, 13), index=2)

    generate_btn = st.button("🚀 キャッシュフロー生成", use_container_width=True, type="primary")

    st.divider()
    debug_mode = st.checkbox("🔍 デバッグモード", value=False)

# セッション状態の初期化
if "html_result" not in st.session_state:
    st.session_state["html_result"] = None
if "cf_data" not in st.session_state:
    st.session_state["cf_data"] = None
if "months" not in st.session_state:
    st.session_state["months"] = None

if generate_btn:
    months = []
    cy, cm = start_year, start_month
    while (cy, cm) <= (end_year, end_month):
        months.append({"year": cy, "month": cm})
        cm += 1
        if cm > 12:
            cm = 1
            cy += 1

    progress = st.progress(0, text="準備中...")
    cf_data  = {}
    verify_data = {}
    debug_logs  = []

    progress.progress(5, text="預金科目を取得中...")
    try:
        deposit_items = get_deposit_accounts(token, company_id)
    except Exception as e:
        st.error(f"預金科目の取得に失敗: {e}")
        st.stop()

    if not deposit_items:
        st.error("預金科目が見つかりません")
        st.stop()

    debug_logs.append(f"**預金科目:** {', '.join(i['name'] for i in deposit_items)}")

    total_steps = len(months) * len(deposit_items)
    step = 0

    for i, mon in enumerate(months):
        key        = str(mon["year"]) + "-" + str(mon["month"])
        last_day   = monthrange(mon["year"], mon["month"])[1]
        start_date = f"{mon['year']}-{mon['month']:02d}-01"
        end_date   = f"{mon['year']}-{mon['month']:02d}-{last_day}"

        all_journals        = []
        freee_end_balance   = 0
        freee_start_balance = 0

        for item in deposit_items:
            step += 1
            pct = int(10 + (step / total_steps) * 80)
            progress.progress(pct, text=f"{mon['year']}年{mon['month']}月 / {item['name']} 取得中...")

            try:
                rows = get_general_ledger(token, company_id, item["id"], start_date, end_date)
                debug_logs.append(f"**{mon['year']}/{mon['month']} {item['name']}:** {len(rows)}件取得")

                if rows and debug_mode:
                    debug_logs.append(f"サンプル行キー: {list(rows[0].keys()) if rows else 'なし'}")

                for row in rows:
                    # freee APIの各種フィールド名に対応
                    debit  = row.get("debit_amount") or row.get("debit") or 0
                    credit = row.get("credit_amount") or row.get("credit") or 0
                    if debit == 0 and credit == 0:
                        continue

                    ob = row.get("opening_balance") or row.get("balance_before") or 0
                    cb = row.get("closing_balance") or row.get("balance_after") or row.get("balance") or 0
                    if ob:
                        freee_start_balance += ob
                    if cb:
                        freee_end_balance = cb

                    all_journals.append({
                        "date":    row.get("date") or row.get("issue_date") or start_date,
                        "amount":  debit - credit,
                        "description": row.get("description") or row.get("memo") or "",
                        "partner": row.get("partner_name") or row.get("partner") or "",
                        "account": row.get("counter_account_name") or row.get("account_item_name") or "",
                        "accountItem": item["name"],
                    })
            except Exception as e:
                st.warning(f"{item['name']}: {e}")
                debug_logs.append(f"⚠ エラー: {e}")

        debug_logs.append(f"**{mon['year']}/{mon['month']} 合計:** {len(all_journals)}件、推定月末残高: {freee_end_balance:,}円")

        if all_journals:
            agg     = aggregate(all_journals)
            opening = freee_start_balance
            closing = freee_end_balance if freee_end_balance else opening + agg["netCF"]
            agg["openingBalance"] = opening
            agg["closingBalance"] = closing
            cf_data[key]     = agg
            verify_data[key] = freee_end_balance
        else:
            cf_data[key]     = {"openingBalance": 0, "closingBalance": 0, "netCF": 0}
            verify_data[key] = 0

    # 月初残高を連鎖
    for i in range(1, len(months)):
        pk = str(months[i-1]["year"]) + "-" + str(months[i-1]["month"])
        ck = str(months[i]["year"])   + "-" + str(months[i]["month"])
        if pk in cf_data and ck in cf_data:
            cf_data[ck]["openingBalance"] = cf_data[pk]["closingBalance"]

    progress.progress(95, text="HTML生成中...")
    account_names = [i["name"] for i in deposit_items]
    html = generate_html(cf_data, selected_name, months, account_names, verify_data)

    # セッションに保存
    st.session_state["html_result"]  = html
    st.session_state["cf_data"]      = cf_data
    st.session_state["months"]       = months
    st.session_state["verify_data"]  = verify_data
    st.session_state["debug_logs"]   = debug_logs
    st.session_state["period_str"]   = f"{start_year}{start_month:02d}-{end_year}{end_month:02d}"

    progress.progress(100, text="完了！")

# ============================================================
# 結果表示（プレビュー）
# ============================================================
if st.session_state.get("html_result"):
    html        = st.session_state["html_result"]
    cf_data     = st.session_state["cf_data"]
    months      = st.session_state["months"]
    verify_data = st.session_state.get("verify_data", {})
    period_str  = st.session_state.get("period_str", "output")

    st.success("✅ キャッシュフロー表が生成されました！内容を確認してからダウンロードしてください。")

    # 残高照合チェック
    has_diff = False
    for mon in months:
        key = str(mon["year"]) + "-" + str(mon["month"])
        cb  = cf_data.get(key, {}).get("closingBalance", 0)
        fb  = verify_data.get(key, 0)
        if abs(cb - fb) > 0 and fb > 0:
            st.warning(f"⚠ {mon['year']}年{mon['month']}月: freee残高との差異 {cb - fb:,}円")
            has_diff = True
    if not has_diff:
        st.info("✅ 全月、freee残高との差異なし")

    # ---- プレビューテーブル ----
    st.subheader("📊 プレビュー")
    col_labels = [f"{m['year']}年{m['month']}月" for m in months]

    def pv(v):
        if v == 0:
            return "—"
        if v < 0:
            return f"▲ {abs(int(v)):,}"
        return f"{int(v):,}"

    rows_def = [
        ("月初繰越残高 (A)",     "openingBalance", "bal"),
        ("売上の入金",           "売上の入金",      "inc"),
        ("収入計 (B)",           "収入計",          "sub_inc"),
        ("原価",                 "原価",            "exp"),
        ("広告宣伝費",           "広告宣伝費",      "exp"),
        ("販管費（外注費含む）", "販管費",          "exp"),
        ("人件費",               "人件費",          "exp"),
        ("税金",                 "税金",            "exp"),
        ("支出計 (C)",           "支出計",          "sub_exp"),
        ("経常収支 (D)=(B)-(C)", "経常収支",        "total"),
        ("借入による収入",       "借入による収入",  "inc"),
        ("貸付の回収",           "貸付の回収",      "inc"),
        ("財務収入計 (E)",       "財務収入計",      "sub_inc"),
        ("貸付による支出",       "貸付による支出",  "exp"),
        ("借入の返済",           "借入の返済",      "exp"),
        ("財務支出計 (F)",       "財務支出計",      "sub_exp"),
        ("財務収支 (G)=(E)-(F)", "財務収支",        "total"),
        ("合計収支 (H)=(D)+(G)", "netCF",           "total"),
        ("次月繰越残高 (I)",     "closingBalance",  "bal"),
    ]

    table_data = {}
    for label, key, _ in rows_def:
        table_data[label] = {col: pv(cf_data.get(str(m["year"]) + "-" + str(m["month"]), {}).get(key, 0)) for col, m in zip(col_labels, months)}

    import pandas as pd
    df = pd.DataFrame(table_data).T
    df.columns = col_labels
    st.dataframe(df, use_container_width=True)

    st.divider()

    # ---- ダウンロードボタン ----
    st.download_button(
        label="⬇️ HTMLをダウンロード",
        data=html.encode("utf-8"),
        file_name=f"cashflow_{period_str}.html",
        mime="text/html",
        use_container_width=True,
        type="primary",
    )

    # ---- デバッグ情報 ----
    if debug_mode:
        if st.session_state.get("all_account_items"):
            with st.expander("📋 freee全勘定科目一覧"):
                import pandas as pd
                df_acct = pd.DataFrame(st.session_state["all_account_items"])
                st.dataframe(df_acct, use_container_width=True)
                st.caption("この一覧から預金として使いたい科目名を確認してください")
        if st.session_state.get("debug_logs"):
            with st.expander("🔍 デバッグログ"):
                for log in st.session_state["debug_logs"]:
                    st.markdown(log)
        if st.session_state.get("last_ledger_response"):
            with st.expander("📦 freee APIレスポンス（最初の取得）"):
                st.json(st.session_state["last_ledger_response"])
