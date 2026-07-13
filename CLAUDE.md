# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Traditional-Chinese, single-page Streamlit business-intelligence dashboard for 嘉楓食品 (Jiafeng Food), a Taiwanese food-distribution company. Managers upload internal Excel/CSV reports and the app parses them, renders KPI/analytics visualizations, calls the Claude API for narrative insights, and can dispatch warning notifications to underperforming salespeople via email or Microsoft Teams.

The entire application lives in one file: `streamlit_app.py` (~1570 lines). There is no package structure, no test suite, and no build step — this is a script-style Streamlit app.

## Commands

Run locally:
```bash
export ANTHROPIC_API_KEY=sk-ant-...
streamlit run streamlit_app.py
```

Install dependencies:
```bash
pip install -r requirements.txt
```

There is no linter, formatter, or test suite configured in this repo. There are no npm/build scripts — it's pure Python.

Deployment target is Streamlit Community Cloud: push to GitHub, set the `ANTHROPIC_API_KEY` secret in the Streamlit Cloud dashboard, and the app is served directly from `streamlit_app.py`. `runtime.txt` / `.python-version` pin Python 3.12; `.streamlit/config.toml` sets the theme and `headless = true` (required for cloud deploy, don't revert it).

## Architecture

`streamlit_app.py` is organized top-to-bottom into clear sections (search for the `# ─── ... ───` banner comments) rather than modules. Reading order roughly matches execution order for a Streamlit script (top-level code runs on every rerun/interaction):

1. **Page config & CSS** — `st.set_page_config`, then a large injected `<style>` block that reskins Streamlit's default widgets (sidebar colors, upload dropzone, buttons). Change styling here, not via Streamlit theme config, for anything beyond basic colors.
2. **Session state defaults** (`_DEFAULTS` dict) — all cross-rerun state lives in `st.session_state`. There are two independent "pipelines" of state: the core 業績分析 (performance) results (`done`, `top_prods`, `top_vends`, `under`, `kpis`, `insights`, ...) and the 通路市場分析 (channel/market) results (all `ch_*` keys). Understanding a feature usually means tracing which `_DEFAULTS` keys it reads/writes.
3. **Report parsers** (`parse_products`, `parse_targets`, `parse_machines`, `parse_channel_file`, `parse_salesperson_data`) — each function hard-codes the row/column layout of a specific Excel export from the company's internal system (fixed header offsets like `df.iloc[4:]`, column indices, and Chinese label strings like `"今年淨銷貨額"` or `"2026達成率"` used as row filters). These are brittle by design: they assume upstream report templates don't change layout. If a parser breaks, the fix is almost always "the source spreadsheet's row/column offsets shifted," not a logic bug.
4. **Analysis helpers** (`top_products`, `top_vendors`, `get_underperformers`, `get_overall_rate`, `build_channel_comparison`, `build_monthly_trend`) — pure pandas transforms over the parsed DataFrames. `get_underperformers` uses a fixed 75% achievement-rate threshold; channel growth/decline thresholds (`±5%`, `±10%`) are hard-coded inline where used, not centralized constants.
5. **Claude AI integration** (`build_prompt`/`call_claude` for performance insights, `build_channel_ai_prompt`/`call_claude_channel` for channel strategy) — builds a Traditional-Chinese prompt from the parsed data, calls `anthropic.Anthropic().messages.create(model="claude-sonnet-4-6", ...)`, and expects the model to return raw JSON (`{"recommendations": [...], "risks": [...]}` or `{"diagnosis": [...], "strategy": [...], "actions": [...]}`). On `JSONDecodeError` it falls back to splitting the raw text into lines. The API key is read from `st.secrets["ANTHROPIC_API_KEY"]` first (Streamlit Cloud), falling back to the `ANTHROPIC_API_KEY` env var (local dev).
6. **Dispatch engine** (`make_dispatch_msg`, `send_via_email`, `send_via_teams`) — formats a warning message from `DISPATCH_MSG_TEMPLATE` and sends it via SMTP and/or a Teams incoming webhook. **These are unconfigured by default**: `SMTP_SERVER`, `SENDER_EMAIL`, `SALESPERSON_EMAILS`, and `MANAGER_EMAILS` near the top of the file are empty and must be filled in per-deployment (SMTP password and Teams webhook URL come from env vars `SMTP_PASSWORD` / `TEAMS_WEBHOOK_URL`). Don't hardcode real credentials into these dict literals — use env vars or `st.secrets` instead, matching the existing pattern.
7. **UI** — sidebar (logo, nav anchor links, live KPI recap, mascot with a rule-based pep-talk keyed off achievement rate via `get_mascot`), then the main page: AI insight banner → upload form → KPI overview → analysis tables (top products/vendors/underperformers) → channel/market analysis section → dispatch engine UI. All layout is hand-built with `st.markdown(..., unsafe_allow_html=True)` and inline CSS-in-Python rather than Streamlit's native components, so most visual tweaks mean editing f-string HTML blocks in place.

Note: there's a large `if False:` block (~line 685–1008) containing an older, now-dead copy of the channel-analysis UI, left in place with a comment noting it was superseded by the inline version further down (section ⑤, ~line 1295). Don't resurrect or duplicate it — the live implementation is the one under `st.markdown('<div id="channel"></div>', ...)`.

## Data flow

Three reports are required (產品銷售年度分析表, 業務目標達成統計, 線上機台數明細表) and two are optional (2025/2026 客戶發展管理報表彙總表, for channel analysis, which also reuses the 業務目標達成統計 file via `parse_salesperson_data` to compute per-salesperson YoY growth). Everything happens in-memory per session — uploaded files are parsed with pandas directly from the `UploadedFile` buffer; nothing is persisted to disk or a database except the two static images in `原始資料/` (`logo.png`, `mascot.png`, loaded and base64-inlined via `img_b64`). Actual `.xlsx`/`.xls`/`.csv` report uploads are explicitly gitignored (`原始資料/*.xlsx` etc.) since they contain real business data — don't commit sample reports into that directory.

## Conventions

- All user-facing strings, comments, and section banners are Traditional Chinese; keep new UI text and prompts consistent with this.
- Section banners (`# ─── 標題 ────...`) delimit logical regions of the file — preserve this style when adding a new section rather than introducing a different comment convention.
- Chart styling (Plotly) consistently uses `plot_bgcolor`/`paper_bgcolor` set to transparent and a shared color-by-severity scheme (red/orange/green thresholds) across bar and line charts — match this when adding new charts.
- The `id="..."` anchor divs (`#ai-insight`, `#upload`, `#kpi`, `#analysis`, `#channel`, `#dispatch`) are targets for the sidebar nav links; if you add/remove/reorder a top-level section, update both the anchor div and the sidebar `_nav_a` links together.
