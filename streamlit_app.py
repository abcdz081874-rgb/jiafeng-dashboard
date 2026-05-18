"""
嘉楓食品 — 業務分析儀表板（Streamlit 版）

啟動方式：
    export ANTHROPIC_API_KEY=sk-ant-...
    streamlit run streamlit_app.py

雲端部署（Streamlit Cloud）：
    1. 推上 GitHub
    2. 在 Streamlit Cloud 設定 ANTHROPIC_API_KEY secret
    3. 直接以網址分享給所有主管
"""
import json
import os
import smtplib
import warnings
from base64 import b64encode
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO, StringIO

import anthropic
import pandas as pd
import requests as http_requests
import streamlit as st

warnings.filterwarnings("ignore")

# ─── 頁面設定（必須第一行）────────────────────────────────────────────────────
st.set_page_config(
    page_title="嘉楓食品 業務分析儀表板",
    page_icon="🍁",
    layout="wide",
    initial_sidebar_state="expanded",
)

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "原始資料")

# ─── 推播設定（請填入貴公司實際值）─────────────────────────────────────────────
SMTP_SERVER       = ""
SMTP_PORT         = 587
SMTP_PASSWORD     = os.environ.get("SMTP_PASSWORD", "")
SENDER_EMAIL      = ""
TEAMS_WEBHOOK_URL = os.environ.get("TEAMS_WEBHOOK_URL", "")
SALESPERSON_EMAILS: dict = {}   # {"A1": "a1@jiafeng.com.tw"}
MANAGER_EMAILS: dict    = {}   # {"北區": "mgr@jiafeng.com.tw"}

DISPATCH_MSG_TEMPLATE = (
    "【嘉楓業務預警】業務 {person} 您好，"
    "您負責的客戶（{region}轄區）目前擁有租賃設備，"
    "但近期原料採購量低於歷史常態"
    "（2026 年度達成率 {rate:.1f}%，"
    "累計實績 {actual:,.0f} 元 / 目標 {target:,.0f} 元）。"
    "請協助於本週安排客訪並確認主因："
    "(A) 受淡旺季影響；"
    "(B) 客戶近期生意下滑，需協助菜單輔導；"
    "(C) 疑似使用競爭對手原料。"
    "拜訪後請回報主管。"
)

# ─── 自訂 CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* 隱藏 Streamlit 預設元素 */
#MainMenu, footer, [data-testid="stDeployButton"] { visibility: hidden; }
[data-testid="stToolbar"] { display: none; }

/* 側邊欄文字顏色 */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown { color: rgba(255,255,255,0.82) !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #fff !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12); }

/* 側邊欄 checkbox */
[data-testid="stSidebar"] .stCheckbox label { color: rgba(255,255,255,0.8) !important; }

/* 上傳區塊 */
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(46,189,126,0.5) !important;
    border-radius: 10px !important;
    background: rgba(46,189,126,0.04) !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #2ebd7e !important;
    background: rgba(46,189,126,0.08) !important;
}

/* 主要按鈕 */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1a8a55, #2ebd7e) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 12px !important;
    box-shadow: 0 4px 14px rgba(46,189,126,0.35) !important;
    transition: opacity .2s !important;
}
.stButton > button[kind="primary"]:hover { opacity: .88 !important; }

/* Expander */
[data-testid="stExpander"] { border-radius: 10px !important; border: 1px solid #dde3ec !important; }

/* 分隔線 */
hr { border-color: #dde3ec; }
</style>
""", unsafe_allow_html=True)


# ─── Session State 初始化 ─────────────────────────────────────────────────────
_DEFAULTS = dict(
    done=False, top_prods=[], top_vends=[], under=[],
    kpis={}, insights=None, ai_warn=None,
    overall_rate=None, total_actual=0.0, total_target=0.0,
    dispatch_msgs=[], dispatch_result=None,
)
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ─── 圖片讀取（base64 嵌入）────────────────────────────────────────────────────
def img_b64(filename: str) -> str:
    path = os.path.join(ASSETS_DIR, filename)
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return b64encode(f.read()).decode()


# ─── 資料讀取 ─────────────────────────────────────────────────────────────────
def read_file(f) -> pd.DataFrame:
    name = getattr(f, "name", "").lower()
    raw = f.read()
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "big5", "gbk"):
            try:
                return pd.read_csv(StringIO(raw.decode(enc)), header=None)
            except Exception:
                continue
        raise ValueError("CSV 編碼無法識別")
    return pd.read_excel(BytesIO(raw), header=None)


# ─── 解析函數 ─────────────────────────────────────────────────────────────────
def parse_products(df: pd.DataFrame):
    data = df.iloc[4:].copy().reset_index(drop=True)
    data[1] = data[1].ffill()
    pm = data[5].astype(str).str.strip() == "今年淨銷貨額"
    products = data[pm][[1, 4, 18]].copy()
    products.columns = ["廠商", "品名", "銷售總和"]
    products["銷售總和"] = pd.to_numeric(products["銷售總和"], errors="coerce").fillna(0)
    products = products.dropna(subset=["品名"])
    vm = (
        pd.to_numeric(data[2], errors="coerce").notna()
        & ~data[1].astype(str).str.contains("總和|合計|小計", na=False)
    )
    vendors = data[vm][[1, 2]].copy()
    vendors.columns = ["廠商", "年銷售額"]
    vendors["年銷售額"] = pd.to_numeric(vendors["年銷售額"], errors="coerce")
    vendors = vendors.dropna().drop_duplicates("廠商")
    return products, vendors


def parse_targets(df: pd.DataFrame):
    data = df.iloc[2:].copy().reset_index(drop=True)
    sub = (
        data[2].astype(str).str.contains("小計|合計|總計", na=False)
        | data[3].astype(str).str.contains("小計|合計|總計", na=False)
    )
    data[1] = data[1].ffill()
    data[3] = data[3].ffill()
    clean = data[~sub].copy()

    def gs(ind):
        rows = clean[clean[4].astype(str).str.strip() == ind]
        return pd.to_numeric(rows.set_index([1, 3])[17], errors="coerce")

    rates, actuals, targets = gs("2026達成率"), gs("2026業績"), gs("2026目標")
    result = []
    for (region, person), rate in rates.items():
        if pd.isna(rate):
            continue
        actual = actuals.get((region, person), 0)
        target = targets.get((region, person), 0)
        result.append(dict(
            person=str(person), region=str(region), rate=float(rate),
            actual=float(actual) if not pd.isna(actual) else 0.0,
            target=float(target) if not pd.isna(target) else 0.0,
        ))
    return result


def parse_machines(df: pd.DataFrame) -> pd.DataFrame:
    header = df.iloc[4]
    total_cols = [i for i, v in enumerate(header) if str(v).strip() in ("總和:", "總和", "小計", "合計")]
    data = df.iloc[5:].copy().reset_index(drop=True)
    mr = data[data[1].notna() & ~data[1].astype(str).str.contains("總和|合計|小計", na=False)].copy()
    if total_cols:
        for c in total_cols:
            mr[c] = pd.to_numeric(mr[c], errors="coerce").fillna(0)
        mr["total"] = mr[total_cols].sum(axis=1)
    else:
        nc = [c for c in data.columns if c not in (0, 1)]
        for c in nc:
            mr[c] = pd.to_numeric(mr[c], errors="coerce").fillna(0)
        mr["total"] = mr[nc].sum(axis=1)
    res = mr[[1, "total"]].copy()
    res.columns = ["機台名稱", "總數"]
    return res[res["總數"] > 0].sort_values("總數", ascending=False)


# ─── 分析 ────────────────────────────────────────────────────────────────────
def top_products(df, n=5):
    return [{"name": r["品名"], "sales": r["銷售總和"]}
            for _, r in df.nlargest(n, "銷售總和").iterrows()]

def top_vendors(df, n=3):
    return [{"name": r["廠商"], "sales": r["年銷售額"]}
            for _, r in df.nlargest(n, "年銷售額").iterrows()]

def get_underperformers(targets, threshold=0.75):
    return sorted([t for t in targets if t["rate"] < threshold], key=lambda x: x["rate"])

def get_overall_rate(targets):
    ta = sum(t["actual"] for t in targets)
    tt = sum(t["target"] for t in targets)
    return (ta / tt if tt > 0 else None), ta, tt


# ─── 吉祥物 ──────────────────────────────────────────────────────────────────
def get_mascot(rate):
    if rate is None:
        return "歡迎！請上傳三份報表，\n我來幫您洞察今日業績！", "#2ebd7e", "#2ebd7e", 0
    p = rate * 100
    if p >= 100:
        return "今天充滿激動的一天！\n大家都是業務奇才呀！🎉", "#f6c90e", "#f6c90e", min(p, 100)
    elif p >= 85:
        return "業績表現亮眼，距離目標\n只差一步！繼續衝刺！💪", "#2ebd7e", "#2ebd7e", p
    elif p >= 70:
        return "快要達標了！再加把勁，\n終點就在前方！🏃", "#4299e1", "#4299e1", p
    elif p >= 50:
        return "今天績效差了點，一定要\n主動出擊客訪！", "#ed8936", "#ed8936", p
    else:
        return "業績需要大幅提升！\n立刻行動，主動拜訪客戶！🔥", "#e53e3e", "#e53e3e", p


# ─── AI ──────────────────────────────────────────────────────────────────────
def build_prompt(top_prods, top_vends, under, machines_df):
    lines = [
        "你是嘉楓食品公司的首席商業分析顧問（台灣食品流通業）。",
        "請根據以下 2026 年度營運數據，以繁體中文給出：",
        "  【經營管理建議】3 點，每點 1-2 句，具體可操作",
        "  【潛在風險提示】3 點，每點 1-2 句，指出隱患與可能衝擊",
        "⚠️ 回應格式必須是純 JSON：",
        '{"recommendations":["建議1","建議2","建議3"],"risks":["風險1","風險2","風險3"]}',
        "", "=== Top 5 銷售產品 ===",
    ]
    for i, p in enumerate(top_prods, 1):
        lines.append(f"  {i}. {p['name']}  銷售額 {p['sales']:,.0f} 元")
    lines.append("\n=== Top 3 廠商 ===")
    for i, v in enumerate(top_vends, 1):
        lines.append(f"  {i}. {v['name']}  年銷售額 {v['sales']:,.0f} 元")
    lines.append("\n=== 業績未達標業務（<75%）===")
    if under:
        for u in under:
            lines.append(f"  - {u['region']} {u['person']}：{u['rate']*100:.1f}%，目標 {u['target']:,.0f}，實績 {u['actual']:,.0f}")
    else:
        lines.append("  無")
    if machines_df is not None and not machines_df.empty:
        total = int(machines_df["總數"].sum())
        lines.append(f"\n=== 線上機台（共 {total} 台）===")
        for _, row in machines_df.head(5).iterrows():
            lines.append(f"  - {row['機台名稱']}：{int(row['總數'])} 台")
    return "\n".join(lines)


def call_claude(prompt: str):
    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        api_key = ""
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return None, "未設定 ANTHROPIC_API_KEY"
    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        s, e = raw.find("{"), raw.rfind("}") + 1
        parsed = json.loads(raw[s:e])
        parsed.setdefault("recommendations", [])
        parsed.setdefault("risks", [])
        return parsed, None
    except json.JSONDecodeError:
        lines = [ln.strip().lstrip("-•1234567890. ") for ln in raw.splitlines() if ln.strip()]
        return {"recommendations": lines[:3], "risks": lines[3:6]}, None
    except Exception as ex:
        return None, str(ex)


# ─── 交辦推播 ─────────────────────────────────────────────────────────────────
def make_dispatch_msg(u: dict) -> str:
    return DISPATCH_MSG_TEMPLATE.format(
        person=u["person"], region=u["region"],
        rate=u["rate"] * 100, actual=u["actual"], target=u["target"],
    )


def send_via_email(under_list):
    results, errors = [], []
    if not SMTP_SERVER or not SENDER_EMAIL:
        errors.append("Email 設定未完成（請在 streamlit_app.py 頂端填入 SMTP_SERVER / SENDER_EMAIL）")
        return results, errors
    try:
        srv = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15)
        srv.ehlo(); srv.starttls()
        if SMTP_PASSWORD:
            srv.login(SENDER_EMAIL, SMTP_PASSWORD)
    except Exception as ex:
        errors.append(f"SMTP 連線失敗：{ex}")
        return results, errors
    for u in under_list:
        body = make_dispatch_msg(u)
        to_e = SALESPERSON_EMAILS.get(u["person"])
        mg_e = MANAGER_EMAILS.get(u["region"])
        recs = [e for e in [to_e, mg_e] if e]
        if not recs:
            errors.append(f"⚠️ {u['region']} {u['person']}：未設定 Email，已略過")
            continue
        mime = MIMEMultipart("alternative")
        mime["From"] = SENDER_EMAIL
        mime["To"] = to_e or ""
        mime["Cc"] = mg_e or ""
        mime["Subject"] = f"【嘉楓業務預警】{u['region']} {u['person']} 業績追蹤通知"
        mime.attach(MIMEText(body, "plain", "utf-8"))
        try:
            srv.sendmail(SENDER_EMAIL, recs, mime.as_string())
            results.append(f"✅ {u['region']} {u['person']} → Email 發送成功（{', '.join(recs)}）")
        except Exception as ex:
            errors.append(f"❌ {u['person']} Email 失敗：{ex}")
    try:
        srv.quit()
    except Exception:
        pass
    return results, errors


def send_via_teams(under_list):
    results, errors = [], []
    if not TEAMS_WEBHOOK_URL:
        errors.append("Teams Webhook URL 未設定（請填入 TEAMS_WEBHOOK_URL 或設定同名環境變數）")
        return results, errors
    lines = [f"⚠️ **嘉楓食品業務預警** — 共 {len(under_list)} 位業務達成率低於 75%\n"]
    for u in under_list:
        lines.append(f"**{u['region']} {u['person']}**（{u['rate']*100:.1f}%）\n" + make_dispatch_msg(u))
    payload = {
        "@type": "MessageCard", "@context": "http://schema.org/extensions",
        "themeColor": "DC2626", "summary": "嘉楓業務預警",
        "sections": [{"activityTitle": "🚀 嘉楓食品主動交辦推播",
                      "activitySubtitle": "業績追蹤系統自動發送",
                      "text": "\n\n---\n\n".join(lines)}],
    }
    try:
        resp = http_requests.post(TEAMS_WEBHOOK_URL, json=payload, timeout=15)
        resp.raise_for_status()
        results.append(f"✅ Teams 訊息已推送至頻道（共 {len(under_list)} 筆）")
    except Exception as ex:
        errors.append(f"❌ Teams 推送失敗：{ex}")
    return results, errors


# ═══════════════════════════════════════════════════════════════════════════════
#  UI
# ═══════════════════════════════════════════════════════════════════════════════

# ─── 側邊欄 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    # Logo
    logo_b64 = img_b64("logo.png")
    if logo_b64:
        st.markdown(
            f'<div style="text-align:center;padding:8px 0 4px">'
            f'<img src="data:image/png;base64,{logo_b64}" style="width:90px;border-radius:12px">'
            f'</div>',
            unsafe_allow_html=True,
        )
    st.markdown(
        '<div style="text-align:center">'
        '<div style="font-size:1.1rem;font-weight:800;color:#fff">嘉楓食品</div>'
        '<div style="font-size:.7rem;color:#2ebd7e;letter-spacing:.08em">BUSINESS INTELLIGENCE</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.divider()

    # 導覽
    st.markdown("**📌 導覽**")
    _nav_a = (
        "display:block;padding:8px 14px;color:rgba(255,255,255,.88);text-decoration:none;"
        "border-radius:8px;font-size:.85rem;font-weight:500;margin-bottom:1px;"
        "transition:background .15s;"
    )
    st.markdown(
        f'<div style="margin:4px -4px">'
        f'<a href="#ai-insight" style="{_nav_a}" onmouseover="this.style.background=\'rgba(46,189,126,.18)\'" onmouseout="this.style.background=\'transparent\'">💡 高階 AI 洞察</a>'
        f'<a href="#upload" style="{_nav_a}" onmouseover="this.style.background=\'rgba(46,189,126,.18)\'" onmouseout="this.style.background=\'transparent\'">📂 上傳報表</a>'
        f'<a href="#kpi" style="{_nav_a}" onmouseover="this.style.background=\'rgba(46,189,126,.18)\'" onmouseout="this.style.background=\'transparent\'">📊 績效總覽</a>'
        f'<a href="#analysis" style="{_nav_a}" onmouseover="this.style.background=\'rgba(46,189,126,.18)\'" onmouseout="this.style.background=\'transparent\'">📈 分析報告</a>'
        f'<a href="#dispatch" style="{_nav_a}" onmouseover="this.style.background=\'rgba(46,189,126,.18)\'" onmouseout="this.style.background=\'transparent\'">🚀 主動交辦推播</a>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.divider()

    # 快速統計（有資料才顯示）
    if st.session_state.done:
        k = st.session_state.kpis
        r = st.session_state.overall_rate
        st.markdown("**📋 本期快報**")
        st.markdown(
            f"<div style='font-size:.82rem;line-height:2'>"
            f"🏭 供應廠商　<b style='color:#fff'>{k.get('total_vendors',0)}</b><br>"
            f"📦 商品品項　<b style='color:#fff'>{k.get('total_products',0)}</b><br>"
            f"🖥️ 線上機台　<b style='color:#fff'>{k.get('total_machines',0)} 台</b><br>"
            f"⚠️ 未達標業務　<b style='color:#ffb347'>{k.get('underperformer_count',0)} 人</b>"
            f"</div>",
            unsafe_allow_html=True,
        )
        if r is not None:
            _, _, rc, dp = get_mascot(r)
            st.markdown(
                f"<div style='font-size:.82rem;margin-top:6px'>"
                f"📈 整體達成率　<b style='color:{rc}'>{r*100:.1f}%</b></div>",
                unsafe_allow_html=True,
            )
        st.divider()

    # 吉祥物
    mascot_msg, bubble_color, _, _ = get_mascot(st.session_state.overall_rate)
    mascot_b64 = img_b64("mascot.png")
    st.markdown(
        f'<div style="text-align:center;margin-top:4px">'
        f'<div style="background:#fff;border:2px solid {bubble_color};border-radius:12px;'
        f'padding:10px 12px;font-size:.78rem;color:#1a202c;line-height:1.6;'
        f'font-weight:500;margin-bottom:10px">{mascot_msg}</div>'
        + (f'<img src="data:image/png;base64,{mascot_b64}" style="width:90px">'
           if mascot_b64 else "🍀")
        + '</div>',
        unsafe_allow_html=True,
    )

    st.divider()
    st.markdown(
        '<div style="font-size:.7rem;color:rgba(255,255,255,.28);text-align:center">'
        '© 2026 嘉楓食品股份有限公司</div>',
        unsafe_allow_html=True,
    )


# ─── 頂部標題列 ───────────────────────────────────────────────────────────────
now = datetime.now()
weekday = ["一", "二", "三", "四", "五", "六", "日"][now.weekday()]
date_str = f"{now.year}/{now.month:02d}/{now.day:02d}　星期{weekday}"

st.markdown(
    f'<div style="background:linear-gradient(120deg,#0d1b3e,#1f4172);color:#fff;'
    f'padding:18px 28px;border-radius:14px;display:flex;align-items:center;'
    f'justify-content:space-between;margin-bottom:24px;box-shadow:0 4px 16px rgba(0,0,0,.2)">'
    f'<span style="font-size:1.3rem;font-weight:800">🍁 嘉楓食品 業務分析儀表板</span>'
    f'<span style="font-size:.82rem;background:rgba(46,189,126,.2);color:#a3f0cc;'
    f'padding:5px 14px;border-radius:20px;border:1px solid rgba(46,189,126,.35)">'
    f'{date_str}</span></div>',
    unsafe_allow_html=True,
)


# ─── ① AI 洞察 ───────────────────────────────────────────────────────────────
st.markdown('<div id="ai-insight"></div>', unsafe_allow_html=True)
if st.session_state.insights:
    ins = st.session_state.insights
    recs = "".join(f'<div style="border-left:3px solid #2ebd7e;padding:9px 14px;margin-bottom:7px;'
                   f'background:rgba(255,255,255,.08);border-radius:0 8px 8px 0;font-size:.88rem;line-height:1.65">'
                   f'{i+1}. {r}</div>' for i, r in enumerate(ins.get("recommendations", [])))
    risks = "".join(f'<div style="border-left:3px solid #ff7043;padding:9px 14px;margin-bottom:7px;'
                    f'background:rgba(255,255,255,.08);border-radius:0 8px 8px 0;font-size:.88rem;line-height:1.65">'
                    f'{i+1}. {r}</div>' for i, r in enumerate(ins.get("risks", [])))
    st.markdown(
        f'<div style="background:linear-gradient(135deg,#0d1b3e,#162447);color:#fff;'
        f'border-radius:16px;padding:24px;border-left:5px solid #2ebd7e;'
        f'box-shadow:0 6px 24px rgba(0,0,0,.25);margin-bottom:24px">'
        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
        f'<span style="background:#1a8a55;color:#fff;font-size:.7rem;font-weight:800;'
        f'letter-spacing:.06em;padding:4px 12px;border-radius:20px">Claude AI</span>'
        f'<span style="font-size:1.05rem;font-weight:700">高階主管每日商業洞察</span></div>'
        f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">'
        f'<div><div style="color:#2ebd7e;font-size:.75rem;font-weight:700;letter-spacing:.06em;margin-bottom:8px">💡 經營管理建議</div>{recs}</div>'
        f'<div><div style="color:#ff9166;font-size:.75rem;font-weight:700;letter-spacing:.06em;margin-bottom:8px">⚠️ 潛在風險提示</div>{risks}</div>'
        f'</div></div>',
        unsafe_allow_html=True,
    )
    if st.session_state.ai_warn:
        st.warning(st.session_state.ai_warn)
else:
    st.info("💡 **高階主管每日商業洞察** — 請上傳三份報表，系統將自動呼叫 Claude AI 產生今日分析。")


# ─── ② 上傳表單 ──────────────────────────────────────────────────────────────
st.markdown('<div id="upload"></div>', unsafe_allow_html=True)
with st.container():
    st.markdown(
        '<div style="background:#fff;border-radius:16px;border:1px solid #dde3ec;'
        'padding:20px 24px 8px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:24px">',
        unsafe_allow_html=True,
    )
    st.markdown("#### 📂 上傳報表（支援 .xlsx / .csv）")
    with st.form("upload_form", clear_on_submit=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown('<div style="font-size:.85rem;font-weight:600;color:#162447;margin-bottom:4px">① 產品銷售年度分析表</div>', unsafe_allow_html=True)
            f1 = st.file_uploader("產品銷售", type=["xlsx", "xls", "csv"], label_visibility="collapsed")
        with c2:
            st.markdown('<div style="font-size:.85rem;font-weight:600;color:#1a8a55;margin-bottom:4px">② 業務目標達成統計</div>', unsafe_allow_html=True)
            f2 = st.file_uploader("業務目標", type=["xlsx", "xls", "csv"], label_visibility="collapsed")
        with c3:
            st.markdown('<div style="font-size:.85rem;font-weight:600;color:#c07000;margin-bottom:4px">③ 線上機台數明細表</div>', unsafe_allow_html=True)
            f3 = st.file_uploader("機台明細", type=["xlsx", "xls", "csv"], label_visibility="collapsed")

        submitted = st.form_submit_button(
            "🚀　上傳並產生 AI 分析",
            type="primary",
            use_container_width=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

# 處理上傳
if submitted:
    if not all([f1, f2, f3]):
        st.error("⚠️ 請同時上傳三份報表再送出")
    else:
        with st.spinner("正在解析報表並呼叫 Claude AI，請稍候（約 15–25 秒）…"):
            try:
                prod_df, vend_df = parse_products(read_file(f1))
                targets          = parse_targets(read_file(f2))
                mach_df          = parse_machines(read_file(f3))

                top_prods = top_products(prod_df)
                top_vends = top_vendors(vend_df)
                under     = get_underperformers(targets)
                rate, ta, tt = get_overall_rate(targets)

                st.session_state.update(
                    done=True,
                    top_prods=top_prods, top_vends=top_vends, under=under,
                    overall_rate=rate, total_actual=ta, total_target=tt,
                    kpis=dict(
                        total_products=len(prod_df),
                        total_vendors=len(vend_df),
                        underperformer_count=len(under),
                        total_machines=int(mach_df["總數"].sum()) if not mach_df.empty else 0,
                    ),
                    dispatch_msgs=[make_dispatch_msg(u) for u in under],
                    dispatch_result=None,
                )

                prompt = build_prompt(top_prods, top_vends, under, mach_df)
                parsed, err = call_claude(prompt)
                st.session_state.insights  = parsed
                st.session_state.ai_warn   = f"AI 分析警告：{err}" if err else None

                st.rerun()

            except Exception as ex:
                st.error(f"資料處理錯誤：{ex}")


# ─── ③ 績效總覽（有資料才顯示）────────────────────────────────────────────────
st.markdown('<div id="kpi"></div>', unsafe_allow_html=True)
if st.session_state.done:
    k  = st.session_state.kpis
    r  = st.session_state.overall_rate
    ta = st.session_state.total_actual
    tt = st.session_state.total_target
    _, _, rc, dp = get_mascot(r)

    # 整體達成率 + KPI 卡
    st.markdown("---")
    st.markdown("#### 📊 績效總覽")

    if r is not None:
        dp_capped = min(dp, 100)
        col_rate, col_k1, col_k2, col_k3, col_k4 = st.columns([2.2, 1, 1, 1, 1])

        with col_rate:
            st.markdown(
                f'<div style="background:linear-gradient(135deg,#162447,#1f4172);color:#fff;'
                f'border-radius:16px;padding:20px 22px;box-shadow:0 4px 16px rgba(0,0,0,.2)">'
                f'<div style="font-size:.75rem;color:rgba(255,255,255,.55);font-weight:700;'
                f'letter-spacing:.04em;margin-bottom:12px">2026 年度整體業績達成率</div>'
                f'<div style="font-size:2.6rem;font-weight:900;color:{rc};line-height:1">{r*100:.1f}%</div>'
                f'<div style="margin-top:10px;height:8px;background:rgba(255,255,255,.15);border-radius:4px;overflow:hidden">'
                f'<div style="height:100%;width:{dp_capped:.1f}%;background:{rc};border-radius:4px;transition:width .6s"></div></div>'
                f'<div style="display:flex;justify-content:space-between;font-size:.76rem;'
                f'color:rgba(255,255,255,.55);margin-top:8px">'
                f'<span>實績 {ta:,.0f}</span><span>目標 {tt:,.0f}</span></div></div>',
                unsafe_allow_html=True,
            )

        def kpi_card_html(icon, val, lbl, grad):
            return (f'<div style="background:{grad};color:#fff;border-radius:14px;'
                    f'padding:18px 16px;box-shadow:0 4px 14px rgba(0,0,0,.15);height:100%">'
                    f'<div style="font-size:1.6rem">{icon}</div>'
                    f'<div style="font-size:1.9rem;font-weight:800;margin-top:4px">{val}</div>'
                    f'<div style="font-size:.76rem;opacity:.82;margin-top:3px">{lbl}</div></div>')

        with col_k1:
            st.markdown(kpi_card_html("🏭", k["total_vendors"], "供應廠商數",
                        "linear-gradient(135deg,#162447,#1f4172)"), unsafe_allow_html=True)
        with col_k2:
            st.markdown(kpi_card_html("📦", k["total_products"], "商品品項數",
                        "linear-gradient(135deg,#1a6b3e,#2ebd7e)"), unsafe_allow_html=True)
        with col_k3:
            st.markdown(kpi_card_html("⚠️", k["underperformer_count"], "業績未達標人數",
                        "linear-gradient(135deg,#7f1d1d,#dc2626)"), unsafe_allow_html=True)
        with col_k4:
            st.markdown(kpi_card_html("🖥️", k["total_machines"], "線上機台總數",
                        "linear-gradient(135deg,#0a4a7a,#1a7ac2)"), unsafe_allow_html=True)
    else:
        c1, c2, c3, c4 = st.columns(4)
        for col, icon, val, lbl, grad in [
            (c1, "🏭", k["total_vendors"],        "供應廠商數",   "linear-gradient(135deg,#162447,#1f4172)"),
            (c2, "📦", k["total_products"],       "商品品項數",   "linear-gradient(135deg,#1a6b3e,#2ebd7e)"),
            (c3, "⚠️", k["underperformer_count"], "業績未達標人數","linear-gradient(135deg,#7f1d1d,#dc2626)"),
            (c4, "🖥️", k["total_machines"],       "線上機台總數", "linear-gradient(135deg,#0a4a7a,#1a7ac2)"),
        ]:
            with col:
                st.markdown(kpi_card_html(icon, val, lbl, grad), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── ④ 分析表格 ─────────────────────────────────────────────────────────
    st.markdown('<div id="analysis"></div>', unsafe_allow_html=True)
    st.markdown("#### 📈 分析報告")
    ca, cb, cc = st.columns(3)

    # Top 5 產品
    RANK_STYLE = {1: "background:#f6c90e;color:#5a3e00", 2: "background:#b0bec5;color:#263238",
                  3: "background:#cc6e2d;color:#fff"}
    with ca:
        st.markdown(
            '<div style="background:#fff;border-radius:14px;border:1px solid #dde3ec;'
            'overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06)">'
            '<div style="padding:13px 18px;font-weight:700;font-size:.9rem;color:#162447;'
            'border-bottom:2px solid #edf1f7">🏆 Top 5 銷售產品</div>',
            unsafe_allow_html=True,
        )
        for i, p in enumerate(st.session_state.top_prods, 1):
            rs = RANK_STYLE.get(i, "background:#edf1f7;color:#4a5568")
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;padding:9px 14px;'
                f'border-bottom:1px solid #f0f4f8;font-size:.82rem">'
                f'<span style="{rs};width:22px;height:22px;border-radius:50%;display:inline-flex;'
                f'align-items:center;justify-content:center;font-weight:800;font-size:.72rem;flex-shrink:0">{i}</span>'
                f'<span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{p["name"]}">{p["name"]}</span>'
                f'<span style="font-weight:600;white-space:nowrap">{p["sales"]:,.0f}</span></div>',
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # Top 3 廠商
    with cb:
        st.markdown(
            '<div style="background:#fff;border-radius:14px;border:1px solid #dde3ec;'
            'overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06)">'
            '<div style="padding:13px 18px;font-weight:700;font-size:.9rem;color:#162447;'
            'border-bottom:2px solid #edf1f7">🏭 Top 3 供應廠商</div>',
            unsafe_allow_html=True,
        )
        for i, v in enumerate(st.session_state.top_vends, 1):
            rs = RANK_STYLE.get(i, "background:#edf1f7;color:#4a5568")
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;padding:9px 14px;'
                f'border-bottom:1px solid #f0f4f8;font-size:.82rem">'
                f'<span style="{rs};width:22px;height:22px;border-radius:50%;display:inline-flex;'
                f'align-items:center;justify-content:center;font-weight:800;font-size:.72rem;flex-shrink:0">{i}</span>'
                f'<span style="flex:1;font-weight:600">{v["name"]}</span>'
                f'<span style="font-weight:600;white-space:nowrap">{v["sales"]:,.0f}</span></div>',
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # 業績未達標
    with cc:
        st.markdown(
            '<div style="background:#fff;border-radius:14px;border:1px solid #dde3ec;'
            'overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06)">'
            '<div style="padding:13px 18px;font-weight:700;font-size:.9rem;color:#162447;'
            'border-bottom:2px solid #edf1f7">⚠️ 業績未達標業務（&lt;75%）</div>',
            unsafe_allow_html=True,
        )
        for u in st.session_state.under:
            pct = u["rate"] * 100
            bar_c = "#dc2626" if pct < 50 else ("#ea580c" if pct < 65 else "#78716c")
            st.markdown(
                f'<div style="padding:9px 14px;border-bottom:1px solid #f0f4f8">'
                f'<div style="display:flex;justify-content:space-between;font-size:.8rem;margin-bottom:3px">'
                f'<span><b>{u["region"]}</b> {u["person"]}</span>'
                f'<span style="color:{bar_c};font-weight:700">{pct:.1f}%</span></div>'
                f'<div style="height:5px;background:#e8edf4;border-radius:3px;overflow:hidden">'
                f'<div style="height:100%;width:{pct:.1f}%;background:{bar_c};border-radius:3px"></div></div>'
                f'<div style="font-size:.73rem;color:#718096;margin-top:2px">實績 {u["actual"]:,.0f}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        if not st.session_state.under:
            st.markdown('<div style="padding:20px;text-align:center;color:#718096">目前無業務低於 75% 🎉</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ─── ⑤ 主動交辦推播引擎 ──────────────────────────────────────────────────
    st.markdown('<div id="dispatch"></div>', unsafe_allow_html=True)
    if st.session_state.under:
        st.markdown("---")
        st.markdown(
            f'<div style="background:linear-gradient(135deg,#7f1d1d,#dc2626);color:#fff;'
            f'border-radius:14px 14px 0 0;padding:14px 22px;font-weight:700;font-size:.95rem">'
            f'🚀 主動交辦推播引擎'
            f'<span style="background:rgba(255,255,255,.2);font-size:.72rem;font-weight:600;'
            f'padding:3px 10px;border-radius:20px;margin-left:10px">'
            f'待通知 {len(st.session_state.under)} 位業務</span></div>',
            unsafe_allow_html=True,
        )

        with st.expander("📋 展開查看全部交辦訊息預覽", expanded=False):
            for u, msg in zip(st.session_state.under, st.session_state.dispatch_msgs):
                pct = u["rate"] * 100
                bc = "#dc2626" if pct < 50 else ("#ea580c" if pct < 65 else "#71717a")
                st.markdown(
                    f'<div style="margin-bottom:12px">'
                    f'<div style="display:flex;gap:6px;align-items:center;margin-bottom:6px">'
                    f'<span style="background:#162447;color:#fff;font-size:.72rem;padding:2px 8px;border-radius:12px">{u["region"]}</span>'
                    f'<span style="background:#1f4172;color:#fff;font-size:.72rem;padding:2px 8px;border-radius:12px">{u["person"]}</span>'
                    f'<span style="background:{bc};color:#fff;font-size:.72rem;padding:2px 8px;border-radius:12px">{pct:.1f}%</span>'
                    f'<span style="font-size:.72rem;color:#718096;margin-left:auto">'
                    f'實績 {u["actual"]:,.0f} / 目標 {u["target"]:,.0f}</span></div>'
                    f'<div style="font-size:.8rem;line-height:1.75;color:#374151;background:#fafafa;'
                    f'border-radius:8px;padding:10px 14px;border-left:3px solid #dc2626">{msg}</div></div>',
                    unsafe_allow_html=True,
                )

        col_chk1, col_chk2 = st.columns(2)
        with col_chk1:
            send_email = st.checkbox("📧 發送 Email 通知（業務本人 + 當區主管）", key="chk_email")
        with col_chk2:
            send_teams = st.checkbox("💬 發送 M365 Teams 頻道訊息", key="chk_teams")

        if st.button("🚀　執行主動交辦", key="btn_dispatch"):
            if not send_email and not send_teams:
                st.warning("⚠️ 請至少選擇一種發送路徑")
            else:
                with st.spinner("正在發送交辦通知…"):
                    results, errors = [], []
                    if send_email:
                        r, e = send_via_email(st.session_state.under)
                        results.extend(r); errors.extend(e)
                    if send_teams:
                        r, e = send_via_teams(st.session_state.under)
                        results.extend(r); errors.extend(e)

                st.markdown(
                    f'<div style="background:#f8fafc;border-radius:10px;padding:16px 20px;'
                    f'border:1px solid #dde3ec;margin-top:12px">'
                    f'<b>發送結果 — 成功 {len(results)} 筆 / 失敗 {len(errors)} 筆</b></div>',
                    unsafe_allow_html=True,
                )
                for r in results:
                    st.success(r)
                for e in errors:
                    st.error(e)
