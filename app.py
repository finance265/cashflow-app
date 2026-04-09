import streamlit as st
import streamlit.components.v1 as components
import requests
import re
import json
import os
import urllib.parse
import pandas as pd
from datetime import date, timedelta
from calendar import monthrange

# ============================================================
# トークンファイル永続化
# ============================================================
TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".freee_tokens.json")

def _load_token_file() -> dict:
    try:
        with open(TOKEN_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _save_token_file(access_token: str, refresh_token: str):
    try:
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            json.dump({"access_token": access_token,
                       "refresh_token": refresh_token,
                       "saved_at": date.today().isoformat()}, f)
    except Exception:
        pass  # Streamlit Cloud など書き込めない環境ではスキップ

def get_oauth_url() -> str:
    """freee OAuth 認証URL を生成"""
    params = {
        "response_type": "code",
        "client_id":     st.secrets.get("FREEE_CLIENT_ID", ""),
        "redirect_uri":  st.secrets.get("REDIRECT_URI", ""),
        "prompt":        "select_company",
    }
    return "https://accounts.secure.freee.co.jp/public_api/authorize?" + urllib.parse.urlencode(params)

def exchange_code_for_tokens(code: str) -> dict:
    """認証コード → アクセストークン＋リフレッシュトークン"""
    redirect_uri = (
        st.session_state.get("_redirect_uri") or
        st.secrets.get("REDIRECT_URI", "")
    )
    res = requests.post(
        "https://accounts.secure.freee.co.jp/public_api/token",
        data={
            "grant_type":    "authorization_code",
            "client_id":     st.secrets.get("FREEE_CLIENT_ID", ""),
            "client_secret": st.secrets.get("FREEE_CLIENT_SECRET", ""),
            "code":          code,
            "redirect_uri":  redirect_uri,
        },
        timeout=15,
    )
    return res.json()

# ── 起動時に OAuth コールバック（?code=XXX）を処理 ──────────────
_qp = st.query_params
if "code" in _qp and not st.session_state.get("_oauth_done"):
    with st.spinner("freeeと認証中..."):
        _data = exchange_code_for_tokens(_qp["code"])
    if _data.get("access_token"):
        st.session_state["stored_access_token"]  = _data["access_token"]
        st.session_state["stored_refresh_token"] = _data.get("refresh_token", "")
        _save_token_file(_data["access_token"], _data.get("refresh_token", ""))
        st.session_state["_oauth_done"] = True
        st.query_params.clear()          # URLから ?code= を除去
        st.rerun()
    else:
        err = _data.get("error_description") or _data.get("error") or str(_data)
        st.error(f"freee認証エラー: {err}")
        st.query_params.clear()

# ── セッション未保持なら token_file から復元 ─────────────────────
if not st.session_state.get("stored_refresh_token"):
    _file = _load_token_file()
    if _file.get("refresh_token"):
        st.session_state["stored_refresh_token"] = _file["refresh_token"]
    if _file.get("access_token") and not st.session_state.get("stored_access_token"):
        st.session_state["stored_access_token"] = _file["access_token"]

st.set_page_config(page_title="キャッシュフロー自動生成", page_icon="🏦", layout="wide")

# ============================================================
# パスワード認証
# ============================================================
def check_auth():
    app_pw = st.secrets.get("APP_PASSWORD", "")
    if not app_pw:
        return True  # パスワード未設定なら認証スキップ
    if st.session_state.get("authenticated"):
        return True
    with st.form("login_form"):
        st.subheader("🔒 ログイン")
        pwd = st.text_input("パスワード", type="password")
        ok  = st.form_submit_button("ログイン", use_container_width=True)
    if ok:
        if pwd == app_pw:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

if not check_auth():
    st.stop()

st.title("🏦 freee キャッシュフロー自動生成")
st.caption("freee会計のデータから自動でキャッシュフロー表を作成します")

FREEE_BASE = "https://api.freee.co.jp/api/1"
TOKEN_URL  = "https://accounts.secure.freee.co.jp/public_api/token"

# ============================================================
# トークン自動更新
# ============================================================
def refresh_access_token():
    client_id     = st.secrets.get("FREEE_CLIENT_ID", "")
    client_secret = st.secrets.get("FREEE_CLIENT_SECRET", "")
    refresh_token = (
        st.session_state.get("stored_refresh_token") or
        st.secrets.get("FREEE_REFRESH_TOKEN", "")
    )
    missing = [k for k, v in [("FREEE_CLIENT_ID", client_id),
                               ("FREEE_CLIENT_SECRET", client_secret),
                               ("FREEE_REFRESH_TOKEN", refresh_token)] if not v]
    if missing:
        st.session_state["_token_error"] = f"Secrets未設定: {', '.join(missing)}"
        return None
    try:
        res = requests.post(TOKEN_URL, data={
            "grant_type":    "refresh_token",
            "client_id":     client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
        }, timeout=10)
        data = res.json()
        if res.status_code != 200:
            err = data.get("error_description") or data.get("error") or res.text
            st.session_state["_token_error"] = f"HTTP {res.status_code}: {err}"
            # 認証エラー(400/401)のときはセッションとファイル両方を破棄
            if res.status_code in (400, 401):
                st.session_state.pop("stored_refresh_token", None)
                st.session_state.pop("stored_access_token", None)
                try:
                    os.remove(TOKEN_FILE)
                except FileNotFoundError:
                    pass
            return None
        if data.get("refresh_token"):
            st.session_state["stored_refresh_token"] = data["refresh_token"]
        st.session_state.pop("_token_error", None)
        return data.get("access_token")
    except Exception as e:
        st.session_state["_token_error"] = f"通信エラー: {e}"
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

def get_account_items(token, company_id):
    """科目マスタ取得: {account_item_id: {"name": str, "category": str}}"""
    d = freee_get("/account_items", token, company_id)
    return {
        item["id"]: {"name": item.get("name", ""), "category": item.get("account_category", "")}
        for item in d.get("account_items", [])
    }

def get_bank_account_item_ids_from_walletables(bank_accounts):
    """walletablesから直接account_item_idを取得（general_ledgers不要）"""
    result = {}
    for w in bank_accounts:
        aid = w.get("account_item_id")
        if aid:
            result[w["name"]] = aid
    return result

def _paginate(path, key, token, company_id, params):
    """ページネーション付きAPIを全件取得。
    freee APIによって total_count の場所が異なるため複数パスで探す。"""
    result = []
    offset = 0
    limit  = 100
    while True:
        d     = freee_get(path, token, company_id, {**params, "offset": offset, "limit": limit})
        items = d.get(key, [])
        result.extend(items)
        # total_count の場所: meta.total_count / total_count / meta.total_entries など
        meta  = d.get("meta") or {}
        total = (meta.get("total_count") or meta.get("total_entries")
                 or d.get("total_count") or 0)
        offset += len(items)
        # total が取れない場合は items が limit 未満になるまで続ける
        if total and offset >= total:
            break
        if len(items) < limit:
            break
    return result

def get_all_transactions(token, company_id, bank_account_item_ids, start_date, end_date, acct_map,
                         debug=False):
    """deals + manual_journals を全件取得。
    Returns: (bank_entries, payable_lookup)
      bank_entries   : 銀行口座が含まれる仕訳のリスト
      payable_lookup : {partner_name: expense_account_name} — 未払金経由払いの費用科目マップ
    """
    params = {"start_issue_date": start_date, "end_issue_date": end_date}
    bank_entries   = []
    payable_lookup = {}  # partner_name → expense_account_name

    PAYABLE_ACCTS = {"未払金", "未払費用", "BD興行未払金", "スクール未払金", "未払消費税等", "未払法人税等"}

    def _resolve_name(aid, name):
        if name:
            return name
        return acct_map.get(aid, {}).get("name", "") if aid else ""

    def _norm_details(details, has_account_name=True):
        out = []
        for det in details:
            aid   = det.get("account_item_id")
            name  = _resolve_name(aid, det.get("account_item_name") if has_account_name else None)
            pname = det.get("partner_name") or det.get("partner_long_name") or ""
            out.append({**det, "account_item_name": name, "partner_name": pname})
        return out

    # 費用系カテゴリ（PLの借方科目のみを payable_lookup に登録する）
    EXPENSE_CATS = {"売上原価", "販売管理費", "営業外費用", "法人税等"}

    def _update_payable_lookup(norm_details, deal_partner=""):
        """未払金 credit / 費用 debit の仕訳から payable_lookup を構築。
        BS科目（売掛金・前払金等）が誤登録されないよう費用科目カテゴリで絞り込む。"""
        # bank が含まれる仕訳はスキップ（別途 bank_entries として処理）
        if any(d.get("account_item_id") in bank_account_item_ids for d in norm_details):
            return
        credits = [d for d in norm_details if d.get("entry_side") == "credit"]
        debits  = [d for d in norm_details if d.get("entry_side") == "debit"]
        # 未払金が credit 側にあるか確認
        payable_credits = [d for d in credits if d.get("account_item_name", "") in PAYABLE_ACCTS]
        if not payable_credits:
            return
        # debit 側の「費用系科目」のみを特定（BS科目・未払金・銀行を除外）
        expense_debits = [
            d for d in debits
            if d.get("account_item_name", "") not in PAYABLE_ACCTS
            and d.get("account_item_id") not in bank_account_item_ids
            and acct_map.get(d.get("account_item_id"), {}).get("category", "") in EXPENSE_CATS
        ]
        if not expense_debits:
            return
        best_debit = max(expense_debits, key=lambda x: x.get("amount", 0) or 0)
        exp_name   = best_debit.get("account_item_name", "")
        if not exp_name:
            return
        # partner 候補を収集して lookup に登録（上書きしない）
        partners = set()
        for pc in payable_credits:
            p = pc.get("partner_name") or deal_partner
            if p:
                partners.add(p)
        p2 = best_debit.get("partner_name") or deal_partner
        if p2:
            partners.add(p2)
        for p in partners:
            if p and p not in payable_lookup:
                payable_lookup[p] = exp_name

    # ---- deals ----
    deals_raw = 0
    deals_hit = 0
    try:
        deals = _paginate("/deals", "deals", token, company_id, params)
        deals_raw = len(deals)
        for deal in deals:
            raw_details  = deal.get("details") or []
            deal_partner = deal.get("partner_name") or deal.get("partner_long_name") or ""
            norm_details = _norm_details(raw_details, has_account_name=True)
            _update_payable_lookup(norm_details, deal_partner)
            if not any(d.get("account_item_id") in bank_account_item_ids for d in raw_details):
                continue
            deals_hit += 1
            bank_entries.append({
                "issue_date":   deal.get("issue_date", ""),
                "partner_name": deal_partner,
                "details":      norm_details,
            })
    except Exception as e:
        st.warning(f"deals取得エラー ({start_date}〜{end_date}): {e}")

    # ---- manual_journals ----
    mj_raw = 0
    mj_hit = 0
    try:
        manuals = _paginate("/manual_journals", "manual_journals", token, company_id, params)
        mj_raw = len(manuals)
        for mj in manuals:
            raw_details  = mj.get("details") or []
            norm_details = _norm_details(raw_details, has_account_name=False)
            _update_payable_lookup(norm_details)
            if not any(d.get("account_item_id") in bank_account_item_ids for d in raw_details):
                continue
            mj_hit += 1
            bank_entries.append({
                "issue_date":   mj.get("issue_date", ""),
                "partner_name": "",
                "details":      norm_details,
            })
    except Exception as e:
        st.warning(f"manual_journals取得エラー ({start_date}〜{end_date}): {e}")

    if debug:
        st.caption(f"　deals: {deals_raw}件 → 銀行含む {deals_hit}件 ／ "
                   f"manual_journals: {mj_raw}件 → 銀行含む {mj_hit}件 ／ "
                   f"payable_lookup: {len(payable_lookup)}件")
        if payable_lookup:
            st.caption(f"　payable_lookup例: { {k: v for k, v in list(payable_lookup.items())[:5]} }")

    return bank_entries, payable_lookup

def get_trial_bs_balances(token, company_id, bank_account_item_ids, year, month, fy_end_month=None):
    """trial_bs から月初・月末残高を取得。
    fy_end_month: 会社の決算月（1-12）。指定があれば正確な fiscal_year を計算する。
    Returns: (opening_balance, closing_balance) or (None, None)"""
    # 正しい fiscal_year を計算する
    if fy_end_month:
        # 決算月から期首月を計算: 決算月=3 → 期首=4, 決算月=12 → 期首=1
        fy_start_month = (fy_end_month % 12) + 1
        # 対象カレンダー月が期首月より前なら前年度
        correct_fy = year if month >= fy_start_month else year - 1
        fy_candidates = [correct_fy]  # 正確な年だけ試す（誤った年で正常応答が返るのを防ぐ）
    else:
        # 決算月不明の場合は両方試す（year-1 を先に試す方が日本企業では多い）
        fy_candidates = [year - 1, year]

    for fy in fy_candidates:
        try:
            d = freee_get("/reports/trial_bs", token, company_id, {
                "fiscal_year":               fy,
                "start_month":               month,
                "end_month":                 month,
                "account_item_display_type": "account_item",
            })
            balances = (d.get("trial_bs") or {}).get("balances", [])
            if not balances:
                continue
            opening = closing = 0
            found = False
            for b in balances:
                aid = b.get("account_item_id")
                if aid in bank_account_item_ids:
                    opening += b.get("opening_balance", 0) or 0
                    closing += b.get("closing_balance", 0) or 0
                    found = True
            if found:
                return opening, closing
        except Exception:
            continue
    return None, None

# ============================================================
# 仕訳から銀行が動いた行を抽出
# ============================================================
def extract_bank_lines(journals, bank_account_item_ids, acct_map=None):
    """
    仕訳帳の全仕訳から銀行口座が含まれる行を抽出してCF行に変換。
    acct_map: {account_item_id: {"name": str, "category": str}} — category付与に使用
    """
    acct_map = acct_map or {}
    results = []
    for j in journals:
        # 仕訳帳APIのレスポンス構造に対応（複数パターン）
        details = j.get("details", []) or j.get("journal_details", [])
        issue_date = j.get("issue_date") or j.get("date") or ""

        if not details:
            # 仕訳帳が行形式の場合（1行=1レコード）
            acct_id = j.get("account_item_id")
            if acct_id in bank_account_item_ids:
                entry_side = j.get("entry_side", "")
                amount     = j.get("amount", 0) or 0
                net        = amount if entry_side == "debit" else -amount
                results.append({
                    "date":        issue_date,
                    "amount":      net,
                    "partner":     j.get("partner_name") or "",
                    "account":     j.get("counter_account_name") or j.get("account_item_name") or "",
                    "description": j.get("description") or j.get("memo") or "",
                    "source":      "journal",
                })
            continue

        # details形式の場合
        bank_lines  = [d for d in details if d.get("account_item_id") in bank_account_item_ids]
        other_lines = [d for d in details if d.get("account_item_id") not in bank_account_item_ids]

        for bl in bank_lines:
            entry_side = bl.get("entry_side", "")
            amount     = bl.get("amount", 0) or 0
            net        = amount if entry_side == "debit" else -amount

            counter_account = ""
            counter_partner = (
                bl.get("partner_name") or bl.get("partner_long_name") or
                j.get("partner_name") or j.get("partner_long_name") or ""
            )

            counter_cat = ""
            if other_lines:
                best = max(other_lines, key=lambda x: x.get("amount", 0))
                counter_account = (
                    best.get("account_item_name") or
                    best.get("account_name") or
                    ""
                )
                # account_category を acct_map から解決
                best_aid = best.get("account_item_id")
                counter_cat = acct_map.get(best_aid, {}).get("category", "")
                if best.get("partner_name") or best.get("partner_long_name"):
                    counter_partner = best.get("partner_name") or best.get("partner_long_name") or ""

            results.append({
                "date":        issue_date,
                "amount":      net,
                "partner":     counter_partner,
                "account":     counter_account,
                "account_cat": counter_cat,   # ← account_category を追加
                "description": bl.get("description") or j.get("description") or "",
                "source":      "journal",
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

# ── account_category（freee共通）→ CF大分類 ─────────────────────────────
# freee の account_category は全事業所共通の体系のため、これをベースにすると汎用性が高い
CATEGORY_CF_MAP = {
    # 収入系 PL
    "売上高":         "売上の入金",   # 売上高, 役務収益, 完成工事高, 売上値引高 等
    "営業外収益":     "売上の入金",   # 受取利息, 雑収入, 受取配当金, 為替差益 等
    "特別利益":       "売上の入金",   # 固定資産売却益, 貸倒引当金戻入益 等

    # 費用系 PL
    "当期商品仕入":   "原価",         # 仕入高, 仕入高_外注費 等
    "棚卸資産":       "原価",         # 商品・仕掛品の増減
    "法人税等":       "税金",
    "法人税等調整額": "税金",
    "特別損失":       "販管費",       # 固定資産除却損, 投資有価証券売却損 等

    # 財務系 BS
    "固定負債":       "_借入",        # 長期借入金の増減（入金=借入, 出金=返済）
    "投資その他の資産": "_貸付",      # 長期貸付金の増減（出金=貸付, 入金=回収）

    # 遡及が必要な BS（銀行仕訳の相手が中間科目のケース）
    "売上債権":       "_要遡及",      # 売掛金・受取手形 → 売上の入金の可能性高
    "他流動資産":     "_要遡及",      # 仮払金・前払費用・立替金・未収入金・短期貸付金 等
    "仕入債務":       "_要遡及",      # 買掛金・工事未払金 → 原価の可能性高
    "他流動負債":     "_要遡及",      # 未払金系（クレカ含む）・仮受金 等

    # 販売管理費は後段の account_name 細分類で処理（ここには含めない）
    # 営業外費用は 支払利息=借入の返済 / 他=販管費 なので後段処理
}

# ── account_name 細分類（販売管理費・営業外費用 内の精緻化に使用）──────────
# ★ CATEGORY_CF_MAP で解決できない科目だけを列挙。新設科目は category で自動対応。
ACCOUNT_NAME_MAP = {
    # 販売管理費カテゴリ内：人件費として扱う科目
    "役員報酬":   "人件費",
    "給与手当":   "人件費",
    "給与":       "人件費",
    "賞与":       "人件費",
    "法定福利費": "人件費",
    "社会保険料": "人件費",
    "労働保険料": "人件費",
    "雇用保険料": "人件費",
    "退職金":     "人件費",

    # 販売管理費カテゴリ内：外注費（個人名なら人件費扱い）
    "外注費":       "_外注",
    "業務委託費":   "_外注",
    "フリーランス": "_外注",

    # 販売管理費カテゴリ内：広告宣伝費
    "広告宣伝費": "広告宣伝費",
    "広告費":     "広告宣伝費",
    "宣伝費":     "広告宣伝費",
    "販促費":     "広告宣伝費",

    # 営業外費用カテゴリ内：利息系は借入の返済
    "支払利息":   "借入の返済",
    "未払利息":   "借入の返済",
    "割引料":     "借入の返済",

    # 法人税等カテゴリ以外の税金系（租税公課は販売管理費カテゴリ）
    "租税公課":   "税金",
    "源泉所得税": "税金",
}

def _resolve_cf(account_name, account_cat, partner, amount, acct_map_entry=None):
    """科目名 + account_category → CF分類を返す汎用関数。
    account_cat: acct_map から取得した account_category 文字列
    acct_map_entry: {"name": str, "category": str}
    """
    # ① account_category ベースの大分類（汎用・新設科目対応）
    cat_cf = CATEGORY_CF_MAP.get(account_cat, "")
    if cat_cf and not cat_cf.startswith("_"):
        # 販売管理費カテゴリ内は account_name で細分類するため別処理
        pass
    elif cat_cf == "_借入":
        return "借入による収入" if amount > 0 else "借入の返済"
    elif cat_cf == "_貸付":
        return "貸付の回収" if amount > 0 else "貸付による支出"
    elif cat_cf == "_要遡及":
        return "_要遡及"

    # ② account_name 細分類（人件費・外注費・広告宣伝費など）
    name_cf = ACCOUNT_NAME_MAP.get(account_name, "")
    if name_cf == "_外注":
        return "人件費" if is_personal_name(partner) else "販管費"
    if name_cf:
        return name_cf

    # ③ account_category の大分類を適用（販売管理費→販管費、営業外収益→売上の入金 等）
    if cat_cf and not cat_cf.startswith("_"):
        return cat_cf

    # ④ フォールバック（category 不明 or 未定義）
    return "売上の入金" if amount > 0 else "販管費"


def _classify_account_name(account_name, partner="", amount=0, acct_map=None):
    """payable_lookup 解決後に使用。科目名 → CF分類を返す。"""
    # acct_map から category を引いて _resolve_cf に委譲
    acct_map = acct_map or {}
    entry = next((v for v in acct_map.values() if v.get("name") == account_name), {})
    account_cat = entry.get("category", "")
    result = _resolve_cf(account_name, account_cat, partner, amount)
    # _要遡及 が返ってきた場合は販管費にフォールバック
    return result if result != "_要遡及" else "販管費"


def classify_row(row, skip_lookup=False, payable_lookup=None, acct_map=None):
    account     = row.get("account", "")
    partner     = row.get("partner", "")
    amount      = row.get("amount", 0)
    account_cat = row.get("account_cat", "")  # extract_bank_lines で付与

    result = _resolve_cf(account, account_cat, partner, amount)

    if result == "_要遡及":
        # payable_lookup で partner → expense_account を解決
        if payable_lookup and partner:
            exp_acct = payable_lookup.get(partner)
            if exp_acct:
                return _classify_account_name(exp_acct, partner, amount, acct_map)
        return None if not skip_lookup else "販管費"

    return result

def get_past_category(partner, lookup_rows):
    """APIコールなし：取得済み全仕訳から同取引先の過去行を検索してカテゴリを推定"""
    if not partner:
        return "販管費"
    cache = st.session_state.setdefault("partner_cache", {})
    if partner in cache:
        return cache[partner]
    for line in lookup_rows:
        if line.get("partner") == partner:
            cat = classify_row(line, skip_lookup=True, acct_map=None)
            if cat and cat != "_要遡及":
                cache[partner] = cat
                return cat
    cache[partner] = "販管費"
    return "販管費"

def aggregate_cf(rows, lookup_rows, payable_lookup=None, acct_map=None):
    cats = {
        "売上の入金": 0, "原価": 0, "広告宣伝費": 0, "販管費": 0,
        "人件費": 0, "税金": 0, "借入による収入": 0, "貸付の回収": 0,
        "貸付による支出": 0, "借入の返済": 0,
    }
    unclassified = []

    for row in rows:
        cat = classify_row(row, payable_lookup=payable_lookup, acct_map=acct_map)
        if cat is None:
            cat = get_past_category(row.get("partner", ""), lookup_rows)
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
        <td class="g-label g-keijo" rowspan="9">経常収支</td>
        <td class="bg-inc"></td><td>売上の入金</td>{cells("売上の入金")}
      </tr>
      <tr class="r-sub-inc"><td class="bg-inc"></td><td>収入計　(B)</td>{cells("収入計")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>原価</td>{cells("原価")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>広告宣伝費</td>{cells("広告宣伝費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>販管費（外注費含む）</td>{cells("販管費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>人件費</td>{cells("人件費")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>税金</td>{cells("税金")}</tr>
      <tr class="r-sub-exp"><td class="bg-exp"></td><td>支出計　(C)</td>{cells("支出計")}</tr>
      <tr class="r-total"><td class="bg-gry"></td><td>経常収支　(D)=(B)-(C)</td>{total_cells("経常収支")}</tr>
      <tr class="r-normal">
        <td class="g-label g-zaim" rowspan="7">財務収支</td>
        <td class="bg-inc"></td><td>借入による収入</td>{cells("借入による収入")}
      </tr>
      <tr class="r-normal"><td class="bg-inc"></td><td>貸付の回収</td>{cells("貸付の回収")}</tr>
      <tr class="r-sub-inc"><td class="bg-inc"></td><td>財務収入計　(E)</td>{cells("財務収入計")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>貸付による支出</td>{cells("貸付による支出")}</tr>
      <tr class="r-normal"><td class="bg-exp"></td><td>借入の返済</td>{cells("借入の返済")}</tr>
      <tr class="r-sub-exp"><td class="bg-exp"></td><td>財務支出計　(F)</td>{cells("財務支出計")}</tr>
      <tr class="r-total"><td class="bg-gry"></td><td>財務収支　(G)=(E)-(F)</td>{total_cells("財務収支")}</tr>
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
    ※ 仕訳帳から銀行口座が動いた仕訳を抽出して集計しています。<br>
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
    token = refresh_access_token()
    if token:
        st.success("✅ freee 連携済み")
    else:
        err_msg = st.session_state.get("_token_error", "")
        if err_msg:
            st.warning(f"⚠ {err_msg}")
        # OAuth 連携ボタン（FREEE_CLIENT_ID が設定済みの場合に表示）
        if st.secrets.get("FREEE_CLIENT_ID"):
            # REDIRECT_URI: Secrets優先、なければサイドバーで入力
            redirect_uri = st.secrets.get("REDIRECT_URI", "")
            if not redirect_uri:
                redirect_uri = st.text_input(
                    "このアプリのURL（REDIRECT_URI）",
                    placeholder="https://xxxx.streamlit.app/",
                    help="freee連携アプリのコールバックURLに登録したURLを入力してください",
                    key="input_redirect_uri",
                )
            if redirect_uri:
                auth_url = (
                    "https://accounts.secure.freee.co.jp/public_api/authorize?"
                    + urllib.parse.urlencode({
                        "response_type": "code",
                        "client_id":     st.secrets.get("FREEE_CLIENT_ID", ""),
                        "redirect_uri":  redirect_uri,
                        "prompt":        "select_company",
                    })
                )
                # redirect_uri をセッションに保持（コールバック時の交換に使用）
                st.session_state["_redirect_uri"] = redirect_uri
                st.link_button("🔗 freeeと連携する", auth_url, use_container_width=True)
                st.caption(f"認証後 `{redirect_uri}` に戻ります。\nfreee連携アプリのコールバックURLと一致させてください。")
            else:
                st.info("上にこのアプリのURLを入力すると連携ボタンが表示されます。")
        else:
            st.caption("FREEE_CLIENT_ID を Secrets に設定するとワンクリック連携が使えます。")
        with st.expander("🔑 アクセストークンを手動入力"):
            token = st.text_input("freee アクセストークン", type="password",
                                  key="manual_token")
            st.caption("freee管理画面 → 連携アプリ → アクセストークンをコピー")

    # 連携リセットボタン（エラー時に保存済みトークンをクリア）
    if st.session_state.get("stored_refresh_token") or os.path.exists(TOKEN_FILE):
        if st.button("🗑 保存済み連携をリセット", use_container_width=True):
            st.session_state.pop("stored_refresh_token", None)
            st.session_state.pop("stored_access_token", None)
            st.session_state.pop("_token_error", None)
            try:
                os.remove(TOKEN_FILE)
            except FileNotFoundError:
                pass
            st.rerun()

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
    st.session_state["partner_cache"] = {}
    months = []
    cy, cm = start_year, start_month
    while (cy, cm) <= (end_year, end_month):
        months.append({"year": cy, "month": cm})
        cm += 1
        if cm > 12:
            cm = 1; cy += 1

    progress = st.progress(0, text="銀行口座を取得中...")

    try:
        bank_accounts = get_bank_walletables(token, company_id)
    except Exception as e:
        st.error(f"口座取得エラー: {e}")
        st.stop()

    if not bank_accounts:
        st.error("銀行口座が見つかりません")
        st.stop()

    bank_names = [b["name"] for b in bank_accounts]

    # walletables から直接 account_item_id を取得（general_ledgers 不要）
    progress.progress(5, text="科目マスタ・口座IDを取得中...")
    bank_id_map = get_bank_account_item_ids_from_walletables(bank_accounts)

    # walletables に account_item_id がない場合のフォールバック（株式会社BACKSTAGE専用。他社では口座名が一致しないため発動しない）
    KNOWN_ACCT_IDS = {
        "みずほ銀行":           745334884,
        "きらぼし銀行":         745335730,
        "東京スター銀行":       746026443,
        "GMOあおぞらネット銀行": 994262667,
    }
    KNOWN_WALLETABLE_IDS = {
        "みずほ銀行":           3329629,
        "きらぼし銀行":         3329633,
        "東京スター銀行":       3332890,
        "GMOあおぞらネット銀行": 4505837,
    }
    for name in bank_names:
        if name not in bank_id_map and name in KNOWN_ACCT_IDS:
            bank_id_map[name] = KNOWN_ACCT_IDS[name]
    # walletable_id もフォールバック補完
    for w in bank_accounts:
        if not w.get("id") and w.get("name") in KNOWN_WALLETABLE_IDS:
            w["id"] = KNOWN_WALLETABLE_IDS[w["name"]]

    bank_account_item_ids = set(bank_id_map.values())

    # 科目マスタ取得（manual_journals の account_item_name 補完に使用）
    try:
        acct_map = get_account_items(token, company_id)
    except Exception as e:
        st.warning(f"科目マスタ取得失敗（処理継続）: {e}")
        acct_map = {}

    # 決算月を取得（trial_bs の fiscal_year を正確に計算するために必要）
    fy_end_month = next(
        (c.get("fiscal_year_end_month") for c in companies if c["id"] == company_id),
        None
    )

    if debug_mode:
        st.info(f"銀行口座IDマップ: {bank_id_map}")
        st.info(f"科目マスタ件数: {len(acct_map)}")
        st.info(f"決算月: {fy_end_month}月")

    cf_data       = {}
    verify_data   = {}
    all_rows_pool = []  # 全月の取得済み行（get_past_category の検索用）

    for i, mon in enumerate(months):
        key      = str(mon["year"]) + "-" + str(mon["month"])
        last_day = monthrange(mon["year"], mon["month"])[1]
        s_date   = f"{mon['year']}-{mon['month']:02d}-01"
        e_date   = f"{mon['year']}-{mon['month']:02d}-{last_day}"
        pct      = int(10 + (i / len(months)) * 70)

        progress.progress(pct, text=f"{mon['year']}年{mon['month']}月 取引取得中...")

        try:
            entries, payable_lookup = get_all_transactions(
                token, company_id, bank_account_item_ids,
                s_date, e_date, acct_map, debug=debug_mode)
            all_rows = extract_bank_lines(entries, bank_account_item_ids, acct_map=acct_map)
            if debug_mode:
                st.info(f"{mon['year']}年{mon['month']}月: entries={len(entries)}, rows={len(all_rows)}, "
                        f"payable_lookup={len(payable_lookup)}件")
            progress.progress(pct + 5, text=f"{mon['year']}年{mon['month']}月 分類中...")
            agg = aggregate_cf(all_rows, all_rows_pool, payable_lookup=payable_lookup, acct_map=acct_map)
            all_rows_pool.extend(all_rows)
        except Exception as e:
            st.warning(f"{mon['year']}年{mon['month']}月 取得エラー: {e}")
            cf_data[key] = {"openingBalance": 0, "closingBalance": 0, "netCF": 0}
            verify_data[key] = 0
            continue

        # trial_bs から残高を取得（walletable_balance API は使わない）
        opening, closing = get_trial_bs_balances(
            token, company_id, bank_account_item_ids, mon["year"], mon["month"],
            fy_end_month=fy_end_month)
        if opening is None:
            st.warning(f"{mon['year']}年{mon['month']}月: trial_bs から残高を取得できませんでした "
                       f"（決算月={fy_end_month}, fiscal_year 候補を確認してください）")
            opening = closing = 0
        agg["openingBalance"] = opening
        agg["closingBalance"] = closing
        cf_data[key]     = agg
        verify_data[key] = closing

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

    # ---- 差異チェック ----
    has_diff = False
    for mon in months:
        key  = str(mon["year"]) + "-" + str(mon["month"])
        cb   = cf_data.get(key, {}).get("closingBalance", 0) or 0
        fb   = verify_data.get(key, 0) or 0
        diff = cb - fb
        if abs(diff) > 0 and fb > 0:
            has_diff = True
            st.error(f"⚠ {mon['year']}年{mon['month']}月: freee残高との差異 **{diff:+,}円**")
            # 差異分析：口座間振替・未分類の可能性が高い仕訳を表示
            rows = cf_data.get(key, {}).get("_all_rows", [])
            uncl = cf_data.get(key, {}).get("_unclassified", [])
            if rows:
                with st.expander(f"🔍 差異原因分析 — {mon['year']}年{mon['month']}月の全仕訳明細（{len(rows)}件）"):
                    df = pd.DataFrame(rows)[["date","account","partner","amount","description"]]
                    df.columns = ["日付","相手科目","取引先","金額","摘要"]
                    st.dataframe(df, use_container_width=True)
                    st.caption(f"合計入出金: {sum(r['amount'] for r in rows):+,}円 ／ "
                               f"次月繰越(計算値): {cb:,}円 ／ freee実残高: {fb:,}円 ／ 差異: {diff:+,}円")
            if uncl:
                with st.expander(f"⚠ 過去参照で分類した仕訳（{len(uncl)}件）"):
                    st.dataframe(pd.DataFrame(uncl), use_container_width=True)

    if not has_diff:
        st.success("✅ 全月、freee残高との差異なし")

    st.info("✅ キャッシュフロー表が生成されました")

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
        on_click=lambda: None,
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
