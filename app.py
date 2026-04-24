import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import pdfplumber
import PyPDF2
import requests as _requests
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(
    page_title="OilCorp | Intelligence Suite",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════
# CUSTOM CSS — Premium Dark Industrial Theme
# ══════════════════════════════════════════════════════════════

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;500;600;700;800&family=Barlow:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
    --bg-base:      #0D1117;
    --bg-surface:   #161B22;
    --bg-card:      #1C2128;
    --bg-hover:     #21262D;
    --border:       #30363D;
    --border-light: #21262D;
    --accent-blue:  #2F81F7;
    --accent-gold:  #D4A843;
    --accent-green: #3FB950;
    --accent-red:   #F85149;
    --accent-amber: #E3B341;
    --text-primary: #E6EDF3;
    --text-secondary: #8B949E;
    --text-muted:   #484F58;
}

html, body, [class*="css"], .stApp {
    background-color: var(--bg-base) !important;
    font-family: 'Barlow', sans-serif !important;
    color: var(--text-primary) !important;
}

#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

.main .block-container {
    padding: 0 2.5rem 3rem 2.5rem !important;
    max-width: 1400px !important;
}

/* ── Sidebar ───────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0D1117 0%, #161B22 100%) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] .stRadio > label {
    color: var(--text-secondary) !important;
    font-size: 0.8rem !important;
    font-family: 'JetBrains Mono', monospace !important;
    letter-spacing: 0.05em !important;
}
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
    color: var(--text-primary) !important;
    font-family: 'Barlow', sans-serif !important;
    font-size: 0.9rem !important;
}

/* ── Hero header ───────────────────────────────────── */
.hero-header {
    background: linear-gradient(135deg, #0f1923 0%, #162032 40%, #0f1923 100%);
    border-bottom: 1px solid var(--border);
    padding: 2.5rem 0 2rem 0;
    margin: 0 -2.5rem 2.5rem -2.5rem;
    position: relative;
    overflow: hidden;
}
.hero-header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: repeating-linear-gradient(90deg, transparent, transparent 80px,
        rgba(47,129,247,0.03) 80px, rgba(47,129,247,0.03) 81px);
}
.hero-header::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--accent-blue), transparent);
}
.hero-inner { padding: 0 2.5rem; position: relative; z-index: 1; }
.hero-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.25em;
    text-transform: uppercase;
    color: var(--accent-blue);
    margin-bottom: 0.6rem;
}
.hero-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2.8rem;
    font-weight: 800;
    letter-spacing: -0.01em;
    line-height: 1;
    color: var(--text-primary);
    margin: 0 0 0.5rem 0;
}
.hero-title span { color: var(--accent-blue); }
.hero-subtitle {
    font-size: 0.95rem;
    color: var(--text-secondary);
    font-weight: 300;
    margin: 0;
}

/* ── Section labels ────────────────────────────────── */
.section-label {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: var(--accent-blue);
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}
.section-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: linear-gradient(90deg, var(--border), transparent);
}

/* ── Upload zone ───────────────────────────────────── */
[data-testid="stFileUploader"] > div {
    background: var(--bg-surface) !important;
    border: 1.5px dashed var(--border) !important;
    border-radius: 10px !important;
    transition: all 0.2s ease !important;
}
[data-testid="stFileUploader"] > div:hover {
    border-color: var(--accent-blue) !important;
    background: var(--bg-card) !important;
}

/* ── Selectboxes ───────────────────────────────────── */
.stSelectbox > div > div {
    background-color: var(--bg-surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 8px !important;
    color: var(--text-primary) !important;
}
.stSelectbox label, .stMultiSelect label {
    color: var(--text-secondary) !important;
    font-size: 0.8rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
    font-family: 'JetBrains Mono', monospace !important;
}
.stMultiSelect > div > div {
    background-color: var(--bg-surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 8px !important;
}
[data-baseweb="tag"] {
    background-color: rgba(47,129,247,0.15) !important;
    border: 1px solid rgba(47,129,247,0.3) !important;
    color: var(--accent-blue) !important;
    border-radius: 4px !important;
    font-size: 0.78rem !important;
}

/* ── Buttons ───────────────────────────────────────── */
.stButton > button {
    background: var(--bg-surface) !important;
    color: var(--text-primary) !important;
    border: 1px solid var(--border) !important;
    border-radius: 8px !important;
    font-family: 'Barlow', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    padding: 0.6rem 1.4rem !important;
    transition: all 0.18s ease !important;
}
.stButton > button:hover {
    background: var(--bg-hover) !important;
    border-color: var(--accent-blue) !important;
    color: var(--accent-blue) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"] {
    background: var(--accent-blue) !important;
    border-color: var(--accent-blue) !important;
    color: #fff !important;
    font-weight: 600 !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #1a3a5c, #1a4a7a) !important;
    color: #7cb9f4 !important;
    border: 1px solid rgba(47,129,247,0.35) !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    width: 100% !important;
}

/* ── Tabs ──────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid var(--border) !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border: none !important;
    color: var(--text-secondary) !important;
    font-family: 'Barlow', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
    padding: 0.8rem 1.4rem !important;
    border-bottom: 2px solid transparent !important;
    transition: all 0.18s ease !important;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--text-primary) !important; }
.stTabs [aria-selected="true"] {
    color: var(--accent-blue) !important;
    border-bottom: 2px solid var(--accent-blue) !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: transparent !important;
    padding: 1.8rem 0 !important;
}

/* ── Expander ──────────────────────────────────────── */
.stExpander {
    background: var(--bg-surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    margin-bottom: 1rem !important;
    overflow: hidden !important;
}
.stExpander > details > summary {
    background: var(--bg-surface) !important;
    color: var(--text-primary) !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
    padding: 1rem 1.2rem !important;
}
.stExpander > details[open] > summary { border-bottom: 1px solid var(--border) !important; }
.stExpander > details > div {
    background: var(--bg-card) !important;
    padding: 1rem 1.2rem !important;
}

/* ── Stat cards ────────────────────────────────────── */
.stat-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1rem;
    margin-bottom: 2rem;
}
.stat-card {
    background: var(--bg-surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.2rem 1.4rem;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s ease;
}
.stat-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
}
.stat-card.blue::before  { background: var(--accent-blue); }
.stat-card.gold::before  { background: var(--accent-gold); }
.stat-card.green::before { background: var(--accent-green); }
.stat-label {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 0.5rem;
}
.stat-value {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2.2rem;
    font-weight: 700;
    line-height: 1;
    color: var(--text-primary);
}
.stat-sub { font-size: 0.78rem; color: var(--text-secondary); margin-top: 0.3rem; }

/* ── Period & info cards ───────────────────────────── */
.period-card {
    background: var(--bg-surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.8rem;
}
.info-panel {
    background: rgba(47,129,247,0.07);
    border: 1px solid rgba(47,129,247,0.2);
    border-radius: 8px;
    padding: 1rem 1.2rem;
    font-size: 0.87rem;
    color: var(--text-secondary);
    margin-bottom: 1.2rem;
    line-height: 1.6;
}
.info-panel strong { color: var(--accent-blue); }

/* ── Tab headings ──────────────────────────────────── */
.tab-heading {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.5rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    color: var(--text-primary);
    margin-bottom: 0.2rem;
}
.tab-subhead {
    font-size: 0.85rem;
    color: var(--text-secondary);
    margin-bottom: 1.6rem;
}

/* ── Export card ───────────────────────────────────── */
.export-card {
    background: var(--bg-surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 2rem;
    text-align: center;
    max-width: 520px;
    margin: 2rem auto;
}
.export-icon { font-size: 3rem; margin-bottom: 1rem; display: block; }
.export-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.4rem;
    font-weight: 700;
    margin-bottom: 0.4rem;
    color: var(--text-primary);
}
.export-desc {
    font-size: 0.85rem;
    color: var(--text-secondary);
    margin-bottom: 1.4rem;
    line-height: 1.5;
}

/* ── Sheet badges ──────────────────────────────────── */
.sheet-badges { display: flex; gap: 0.5rem; flex-wrap: wrap; justify-content: center; margin-bottom: 1.4rem; }
.sheet-badge {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 0.3rem 0.8rem;
    font-size: 0.75rem;
    color: var(--text-secondary);
    font-family: 'JetBrains Mono', monospace;
}
.sheet-badge.active {
    background: rgba(63,185,80,0.1);
    border-color: rgba(63,185,80,0.3);
    color: var(--accent-green);
}

/* ── Depot manager ─────────────────────────────────── */
.depot-manager {
    background: var(--bg-surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
    margin-bottom: 1.2rem;
}
.depot-manager-header {
    background: var(--bg-card);
    border-bottom: 1px solid var(--border);
    padding: 0.9rem 1.2rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.depot-manager-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--text-primary);
}
.depot-count-badge {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    background: rgba(47,129,247,0.15);
    color: var(--accent-blue);
    border: 1px solid rgba(47,129,247,0.25);
    border-radius: 20px;
    padding: 0.15rem 0.6rem;
}
.depot-list-body { padding: 0.8rem 1.2rem 1rem 1.2rem; }
.depot-row {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    padding: 0.45rem 0;
    border-bottom: 1px solid var(--border-light);
    font-size: 0.87rem;
    color: var(--text-primary);
}
.depot-row:last-child { border-bottom: none; }
.depot-row-index { font-family: 'JetBrains Mono', monospace; font-size: 0.65rem; color: var(--text-muted); min-width: 20px; text-align: right; }
.depot-check { color: var(--accent-green); font-size: 0.85rem; }
.depot-name { flex: 1; }
.add-depot-title {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 0.6rem;
}

/* ── Query counter ─────────────────────────────────── */
.query-counter {
    background: var(--bg-surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.9rem 1.2rem;
    display: flex;
    align-items: center;
    gap: 0.8rem;
    margin-bottom: 1rem;
    font-size: 0.85rem;
    color: var(--text-secondary);
}
.query-number {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.4rem;
    font-weight: 700;
    color: var(--accent-blue);
    line-height: 1;
}

/* ── Log console ───────────────────────────────────── */
.log-console {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    background: #0d1117;
    border: 1px solid var(--border);
    color: #d4d4d4;
    padding: 12px 16px;
    border-radius: 8px;
    max-height: 200px;
    overflow-y: auto;
    line-height: 1.7;
}

/* ── Custom divider ────────────────────────────────── */
.custom-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--border), transparent);
    margin: 1.5rem 0;
}

/* ── Alerts ────────────────────────────────────────── */
.stSuccess { background: rgba(63,185,80,0.08) !important; border-left: 3px solid var(--accent-green) !important; }
.stWarning { background: rgba(227,179,65,0.08) !important; border-left: 3px solid var(--accent-amber) !important; }
.stError   { background: rgba(248,81,73,0.08) !important; border-left: 3px solid var(--accent-red) !important; }
.stInfo    { background: rgba(47,129,247,0.08) !important; border-left: 3px solid var(--accent-blue) !important; }

/* ── Progress ──────────────────────────────────────── */
.stProgress > div > div { background: var(--bg-surface) !important; border-radius: 100px !important; height: 6px !important; }
.stProgress > div > div > div { background: linear-gradient(90deg, var(--accent-blue), #79b8ff) !important; border-radius: 100px !important; }

/* ── Scrollbar ─────────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg-base); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

/* ── Daily order status pills ──────────────────────── */
.status-pill {
    display: inline-block;
    padding: 0.2rem 0.7rem;
    border-radius: 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.status-pill.released  { background: rgba(63,185,80,0.15); color: var(--accent-green); border: 1px solid rgba(63,185,80,0.3); }
.status-pill.submitted { background: rgba(47,129,247,0.15); color: var(--accent-blue);  border: 1px solid rgba(47,129,247,0.3); }
.status-pill.pending   { background: rgba(227,179,65,0.15); color: var(--accent-amber); border: 1px solid rgba(227,179,65,0.3); }

/* ── Daily metrics row ─────────────────────────────── */
.daily-metric-row {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 0.9rem;
    margin-bottom: 1.6rem;
}
.daily-metric {
    background: var(--bg-surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1rem 1.2rem;
}
.daily-metric-label {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.6rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 0.35rem;
}
.daily-metric-value {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: var(--text-primary);
    line-height: 1;
}
.daily-metric-sub { font-size: 0.75rem; color: var(--text-secondary); margin-top: 0.25rem; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# STYLING HELPERS
# ══════════════════════════════════════════════════════════════
DARK_BLUE  = "1F3864"
MED_BLUE   = "2E75B6"
LIGHT_BLUE = "BDD7EE"
YELLOW     = "FFD966"
ORANGE     = "F4B942"
GREEN      = "E2EFDA"
HEADER_FG  = "FFFFFF"

def _border(style="thin"):
    s = Side(border_style=style, color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=11):
    return Font(bold=bold, color=color, size=size, name="Arial")

def _align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ══════════════════════════════════════════════════════════════
# ENVIRONMENT CONFIG
# ══════════════════════════════════════════════════════════════
OILCORP_USER_ID = os.getenv("OILCORP_USER_ID", "123293")
OILCORP_BDC_ID  = os.getenv("OILCORP_BDC_ID",  "20900")
NPA_COMPANY_ID  = os.getenv("NPA_COMPANY_ID",   "1")
NPA_APP_ID      = os.getenv("NPA_APP_ID",       "3")

STOCK_TXN_URL = os.getenv(
    "NPA_STOCK_TRANSACTION_URL",
    "https://iml.npa-enterprise.com/NewNPA/home/CreateStockTransactionReport"
)
DAILY_ORDERS_URL = os.getenv(
    "NPA_DAILY_ORDERS_URL",
    "https://iml.npa-enterprise.com/NewNPA/home/CreateDailyOrderReport"
)

PRODUCT_MAP = {
    "PMS":    int(os.getenv("PRODUCT_PREMIUM_ID", "12")),
    "GASOIL": int(os.getenv("PRODUCT_GASOIL_ID",  "14")),
    "LPG":    int(os.getenv("PRODUCT_LPG_ID",     "28")),
}

MONTHS = {
    1:"January", 2:"February", 3:"March",    4:"April",
    5:"May",     6:"June",     7:"July",      8:"August",
    9:"September",10:"October",11:"November",12:"December"
}


def _load_depot_map() -> dict:
    depot_map = {}
    for key, value in os.environ.items():
        if not key.startswith("DEPOT_"):
            continue
        raw = key[6:]
        name = raw.replace("_", " ").strip()
        # Apply canonical name fixes
        fixes = {
            "GHANA OIL COLTD TAKORADI":              "GHANA OIL CO.LTD, TAKORADI",
            "GOIL LPG BOTTLING PLANT TEMA":          "GOIL LPG BOTTLING PLANT -TEMA",
            "GOIL LPG BOTTLING PLANT KUMASI":        "GOIL LPG BOTTLING PLANT- KUMASI",
            "NEWGAS CYLINDER BOTTLING LIMITED TEMA": "NEWGAS CYLINDER BOTTLING LIMITED-TEMA",
            "CHASE PETROLEUM TEMA":                  "CHASE PETROLEUM - TEMA",
            "TEMA FUEL COMPANY TFC":                 "TEMA FUEL COMPANY (TFC)",
            "TEMA MULTI PRODUCTS TMPT":              "TEMA MULTI PRODUCTS (TMPT)",
            "TEMA OIL REFINERY TOR":                 "TEMA OIL REFINERY (TOR)",
            "GHANA OIL COMPANY LTD SEKONDI NAVAL BASE": "GHANA OIL COMPANY LTD (SEKONDI NAVAL BASE)",
            "GHANSTOCK LIMITED TAKORADI":            "GHANSTOCK LIMITED (TAKORADI)",
            "SENTUO OIL REFINERY TEMA":              "SENTUO OIL REFINERY - TEMA",
        }
        name = fixes.get(name, name)
        try:
            depot_map[name] = int(value)
        except ValueError:
            pass
    return depot_map

DEPOT_MAP = _load_depot_map()

_HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/pdf,text/html,*/*;q=0.8",
}

def _fetch_pdf(url: str, params: dict, timeout: int = 90) -> bytes | None:
    try:
        r = _requests.get(url, params=params, headers=_HTTP_HEADERS, timeout=timeout)
        r.raise_for_status()
        return r.content if r.content[:4] == b"%PDF" else None
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════
# STOCK TRANSACTION PDF PARSER
# ══════════════════════════════════════════════════════════════
DESCRIPTIONS = sorted([
    "Balance b/fwd", "Stock Take", "Sale",
    "Custody Transfer In", "Custody Transfer Out", "Product Outturn",
], key=len, reverse=True)

SKIP_PFX = (
    "national petroleum authority", "stock transaction report",
    "bdc :", "depot :", "product :", "printed by", "printed on",
    "date trans #", "actual stock balance", "stock commitments",
    "available stock balance", "last stock update", "i.t.s from",
)

def _skip(line):
    lo = line.strip().lower()
    return lo.startswith(SKIP_PFX) or bool(re.match(r"^\d{1,2}\s+\w+,\s+\d{4}", line.strip()))

def _pnum(s):
    s = s.strip()
    neg = s.startswith("(") and s.endswith(")")
    try:
        v = int(s.strip("()").replace(",", ""))
        return -v if neg else v
    except ValueError:
        return None

_DATE_LINE_RE = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(.*)")
_TAIL_RE = re.compile(r"(\([\d,]+\)|[\d,]+)\s+(\([\d,]+\)|[\d,]+)\s*$")

def _parse_any_date_line(line: str) -> dict | None:
    line = line.strip()
    m = _DATE_LINE_RE.match(line)
    if not m:
        return None
    date  = m.group(1)
    trans = m.group(2)
    rest  = m.group(3).strip()
    for search_str in (rest, trans + " " + rest):
        tail = _TAIL_RE.search(search_str)
        if not tail:
            continue
        vol = _pnum(tail.group(1))
        bal = _pnum(tail.group(2))
        if vol is None or bal is None:
            continue
        middle = search_str[: tail.start()].strip()
        for d in DESCRIPTIONS:
            if middle.lower().startswith(d.lower()):
                acct       = middle[len(d):].strip()
                real_trans = "" if d == "Balance b/fwd" else trans
                return {
                    "Date":        date,
                    "Trans #":     real_trans,
                    "Description": d,
                    "Account":     acct,
                    "Volume":      vol,
                    "Balance":     bal,
                }
    return None

def parse_stock_transaction_pdf(pdf_bytes: bytes) -> list:
    records = []
    seen    = set()
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for raw in text.split("\n"):
                    line = raw.strip()
                    if not line or _skip(line):
                        continue
                    row = _parse_any_date_line(line)
                    if row:
                        key = (row["Date"], row["Trans #"], row["Description"], row["Volume"])
                        if key not in seen:
                            seen.add(key)
                            records.append(row)
    except Exception:
        pass
    return records


# ══════════════════════════════════════════════════════════════
# OPENING / CLOSING STOCK FETCHER
# ══════════════════════════════════════════════════════════════
def fetch_oilcorp_stock_balances(year, month, depot_name, depot_id, product, product_id):
    import calendar
    last_day  = calendar.monthrange(year, month)[1]
    start_str = f"{month:02d}/01/{year}"
    end_str   = f"{month:02d}/{last_day:02d}/{year}"
    params = {
        "lngProductId": product_id,
        "lngBDCId":     OILCORP_BDC_ID,
        "lngDepotId":   depot_id,
        "dtpStartDate": start_str,
        "dtpEndDate":   end_str,
        "lngUserId":    OILCORP_USER_ID,
    }
    pdf_bytes = _fetch_pdf(STOCK_TXN_URL, params)
    if not pdf_bytes:
        return {"opening": None, "closing": None, "records": [], "error": "No PDF returned"}
    records = parse_stock_transaction_pdf(pdf_bytes)
    if not records:
        return {"opening": None, "closing": None, "records": [], "error": "No transactions parsed"}

    bfwd_balance = None
    bfwd_index   = None
    first_stock_take = None
    for i, r in enumerate(records):
        if r["Description"] == "Balance b/fwd" and bfwd_balance is None:
            bfwd_balance = float(r["Balance"])
            bfwd_index   = i
        elif (r["Description"] == "Stock Take" and bfwd_index is not None
              and i > bfwd_index and first_stock_take is None):
            first_stock_take = float(r["Balance"])

    opening = first_stock_take if first_stock_take is not None else bfwd_balance
    closing = None
    for r in reversed(records):
        if r["Description"] != "Balance b/fwd":
            closing = float(r["Balance"])
            break
    if closing is None and records:
        closing = float(records[-1]["Balance"])

    return {"opening": opening, "closing": closing, "records": records, "error": None}


# ══════════════════════════════════════════════════════════════
# DAILY ORDERS PARSER — OilCorp only
# ══════════════════════════════════════════════════════════════

_DAILY_LOADED_KW = {"Released", "Submitted"}
_DAILY_HEADER_KW = [
    "ORDER REPORT", "National Petroleum Authority", "ORDER NUMBER",
    "ORDER DATE", "ORDER STATUS", "BDC:", "Total for :", "Printed By :",
    "Page ", "BRV NUMBER", "VOLUME",
]
_PRODUCT_MAP_DAILY = {"AGO": "GASOIL", "PMS": "PREMIUM", "LPG": "LPG"}


def _detect_product_daily(line: str) -> str:
    raw = "AGO" if "AGO" in line else "LPG" if "LPG" in line else "PMS"
    return _PRODUCT_MAP_DAILY.get(raw, raw)


def _parse_daily_loaded_line(line: str, product: str, depot: str) -> dict | None:
    tokens = line.split()
    if len(tokens) < 6:
        return None
    rel_idx = next((i for i, t in enumerate(tokens) if t in _DAILY_LOADED_KW), None)
    if rel_idx is None or rel_idx < 2:
        return None
    try:
        date_tok  = tokens[0]
        order_num = tokens[1]
        volume    = float(tokens[-1].replace(",", ""))
        price     = float(tokens[-2].replace(",", ""))
        brv       = tokens[-3]
        company   = " ".join(tokens[rel_idx + 1:-3]).strip()
        try:
            date_str = datetime.strptime(date_tok, "%d-%b-%Y").strftime("%Y-%m-%d")
        except Exception:
            date_str = date_tok
        return {
            "Date":         date_str,
            "Order Number": order_num,
            "OMC":          company,
            "BRV":          brv,
            "Product":      product,
            "Depot":        depot,
            "Quantity (L)": volume,
            "Price (₵/L)":  price,
            "Status":       [t for t in tokens if t in _DAILY_LOADED_KW][0],
        }
    except Exception:
        return None


def extract_oilcorp_daily_orders(pdf_bytes: bytes) -> pd.DataFrame:
    """Parse a Daily Order PDF — keeps only OilCorp Energia as BDC."""
    rows      = []
    cur_depot = "Unknown"
    cur_prod  = "PREMIUM"
    cur_bdc   = ""

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=2, y_tolerance=2)
                if not text:
                    continue
                for line in text.split("\n"):
                    cl = line.strip()
                    if not cl:
                        continue
                    # Parse metadata lines
                    if "DEPOT:" in cl:
                        m = re.search(r"DEPOT:([^-\n]+)", cl)
                        if m:
                            cur_depot = m.group(1).strip()
                        continue
                    if "BDC:" in cl:
                        m = re.search(r"BDC:([^\n]+)", cl)
                        if m:
                            cur_bdc = m.group(1).strip()
                        continue
                    if "PRODUCT" in cl.upper():
                        cur_prod = _detect_product_daily(cl)
                        continue
                    if any(h in cl for h in _DAILY_HEADER_KW):
                        continue
                    if any(kw in cl for kw in _DAILY_LOADED_KW):
                        row = _parse_daily_loaded_line(cl, cur_prod, cur_depot)
                        if row:
                            row["BDC"] = cur_bdc
                            rows.append(row)
    except Exception:
        pass

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    # Keep only OilCorp Energia rows
    oilcorp_mask = df["BDC"].str.upper().str.contains("OILCORP", na=False)
    df = df[oilcorp_mask].copy()
    df = df.drop_duplicates(subset=["Order Number", "BRV", "Date", "Product"])

    try:
        df["_ds"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.sort_values("_ds").drop(columns=["_ds"])
    except Exception:
        pass

    return df.reset_index(drop=True)


def fetch_oilcorp_daily_orders(start_date: datetime, end_date: datetime) -> pd.DataFrame:
    """Fetch OilCorp daily orders for a date range from the NPA API."""
    start_str = start_date.strftime("%m/%d/%Y")
    end_str   = end_date.strftime("%m/%d/%Y")

    params = {
        "lngCompanyId":    NPA_COMPANY_ID,
        "szITSfromPersol": "persol",
        "strGroupBy":      "DEPOT",
        "strGroupBy1":     "",
        "strQuery1":       "",
        "strQuery2":       start_str,
        "strQuery3":       end_str,
        "strQuery4":       "",
        "strPicHeight":    "1",
        "strPicWeight":    "1",
        "intPeriodID":     "-1",
        "iUserId":         OILCORP_USER_ID,
        "iAppId":          NPA_APP_ID,
    }
    pdf_bytes = _fetch_pdf(DAILY_ORDERS_URL, params)
    if not pdf_bytes:
        return pd.DataFrame()
    return extract_oilcorp_daily_orders(pdf_bytes)


# ══════════════════════════════════════════════════════════════
# ORDER REQUEST DATA LOADER
# ══════════════════════════════════════════════════════════════
def load_order_data(file):
    df = pd.read_excel(file, sheet_name="ORDER REQUEST", header=8)
    df = df[["DATE", "Name of OMC", "Product", "Depot", "Quantity", "Comments"]].copy()
    df.columns = ["DATE", "OMC", "Product", "Depot", "Quantity", "Comments"]
    df = df.dropna(subset=["DATE", "OMC", "Depot", "Quantity"])
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df = df.dropna(subset=["DATE"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    return df


# ══════════════════════════════════════════════════════════════
# P1 / SUMMARY BUILDERS
# ══════════════════════════════════════════════════════════════
def build_p1_tables(df):
    tables = []
    depots = sorted(df["Depot"].dropna().unique())
    for depot in depots:
        for window, (d_lo, d_hi) in [("W1", (1, 15)), ("W2", (16, 31))]:
            mask = (df["Depot"] == depot) & (df["DATE"].dt.day >= d_lo) & (df["DATE"].dt.day <= d_hi)
            sub = df[mask]
            if sub.empty:
                continue
            for product in sorted(sub["Product"].dropna().unique()):
                psub = sub[sub["Product"] == product]
                if psub.empty:
                    continue
                pivot = psub.groupby("OMC")["Quantity"].sum().reset_index()
                pivot = pivot.sort_values("OMC")
                pivot.columns = ["OMC", "Quantity"]
                pivot["TOTAL"] = pivot["Quantity"]
                tables.append({
                    "title": f"{depot} {product} {window}",
                    "depot": depot, "product": product, "window": window,
                    "data": pivot,
                })
    return tables


def build_summary(df):
    df2 = df.copy()
    df2["DepotGroup"] = df2["Depot"].apply(
        lambda x: "BOST" if str(x).upper().startswith("BOST") else x
    )
    pivot = df2.groupby(["DepotGroup", "Product"])["Quantity"].sum().unstack(fill_value=0)
    for col in ["AGO", "PMS", "LPG"]:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot[["AGO", "PMS", "LPG"]]
    pivot["GRAND TOTAL"] = pivot.sum(axis=1)
    pivot.loc["GRAND TOTAL"] = pivot.sum()
    pivot = pivot.reset_index()
    pivot.columns.name = None
    return pivot


# ══════════════════════════════════════════════════════════════
# EXCEL WRITERS
# ══════════════════════════════════════════════════════════════
def write_p1_sheet(ws, tables):
    row = 1
    for tbl in tables:
        data  = tbl["data"]
        title = tbl["title"]
        ws.cell(row, 1, title).font = _font(bold=True, color=HEADER_FG, size=12)
        ws.cell(row, 1).fill = _fill(DARK_BLUE)
        ws.cell(row, 1).alignment = _align()
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        row += 1
        for ci, label in enumerate(["OMC", "QUANTITY", "TOTAL"], 1):
            c = ws.cell(row, ci, label)
            c.font = _font(bold=True, color=HEADER_FG)
            c.fill = _fill(MED_BLUE)
            c.alignment = _align()
            c.border = _border()
        row += 1
        start_data = row
        for _, r in data.iterrows():
            ws.cell(row, 1, r["OMC"]).alignment = _align(h="left")
            ws.cell(row, 1).border = _border()
            for ci, val in enumerate([r["Quantity"], r["TOTAL"]], 2):
                c = ws.cell(row, ci, val)
                c.number_format = "#,##0"
                c.alignment = _align()
                c.border = _border()
            row += 1
        gt_cell = ws.cell(row, 1, "GRAND TOTAL")
        gt_cell.font = _font(bold=True)
        gt_cell.fill = _fill(YELLOW)
        gt_cell.border = _border()
        gt_cell.alignment = _align(h="left")
        for ci in [2, 3]:
            col_letter = get_column_letter(ci)
            c = ws.cell(row, ci, f"=SUM({col_letter}{start_data}:{col_letter}{row-1})")
            c.font = _font(bold=True)
            c.fill = _fill(YELLOW)
            c.number_format = "#,##0"
            c.alignment = _align()
            c.border = _border()
        row += 3
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18


def write_summary_sheet(ws, summary_df):
    ws.cell(1, 1, "LOADING SUMMARY").font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    headers = ["DEPOT", "AGO", "PMS", "LPG", "GRAND TOTAL"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(2, ci, h)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(MED_BLUE)
        c.alignment = _align()
        c.border = _border()
    for ri, row_data in summary_df.iterrows():
        excel_row = ri + 3
        is_total  = str(row_data.iloc[0]) == "GRAND TOTAL"
        fill      = _fill(ORANGE) if is_total else _fill(GREEN) if ri % 2 == 0 else PatternFill()
        for ci, val in enumerate(row_data, 1):
            c = ws.cell(excel_row, ci, val)
            c.border = _border()
            c.alignment = _align(h="left" if ci == 1 else "center")
            c.fill = fill
            if ci > 1:
                c.number_format = "#,##0"
            if is_total:
                c.font = _font(bold=True)
    for col, width in zip(["A", "B", "C", "D", "E"], [20, 14, 14, 10, 16]):
        ws.column_dimensions[col].width = width


def _cell(ws, row, col, value, bold=False, bg=None, fg="000000",
          size=11, h_align="center", num_fmt=None, border=True):
    c = ws.cell(row, col, value)
    c.font = _font(bold=bold, color=fg, size=size)
    c.alignment = _align(h=h_align)
    if bg:
        c.fill = _fill(bg)
    if border:
        c.border = _border()
    if num_fmt:
        c.number_format = num_fmt
    return c


def write_stock_balance_sheet(ws, balance_data, sheet_type, month_label):
    PRODUCTS  = ["PMS", "GASOIL", "LPG"]
    title_text = f"OILCORP ENERGIA LIMITED — {sheet_type} STOCK BALANCE ({month_label})"
    num_cols   = len(PRODUCTS) + 3
    ws.cell(1, 1, title_text).font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    subtitle = (
        "Balance carried forward at start of month (Stock Take if available, else b/fwd)"
        if sheet_type == "OPENING"
        else "Last recorded running balance at end of month"
    )
    ws.cell(2, 1, subtitle).font = _font(bold=False, color="595959", size=10)
    ws.cell(2, 1).fill = _fill(LIGHT_BLUE)
    ws.cell(2, 1).alignment = _align()
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    headers = ["DEPOT"] + [f"{p} (LT/KG)" for p in PRODUCTS] + ["GRAND TOTAL (LT/KG)", "STATUS"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(3, ci, h)
        c.font = _font(bold=True, color=HEADER_FG, size=11)
        c.fill = _fill(MED_BLUE)
        c.alignment = _align()
        c.border = _border()
    seen_depots = []
    depot_data: dict = {}
    for entry in balance_data:
        depot   = entry["depot"]
        product = entry["product"]
        if depot not in depot_data:
            depot_data[depot] = {}
            seen_depots.append(depot)
        depot_data[depot][product] = {"balance": entry["balance"], "error": entry["error"]}
    row = 4
    product_totals = {p: 0.0 for p in PRODUCTS}
    for depot in seen_depots:
        alt_fill  = "F2F2F2" if row % 2 == 0 else None
        prod_info = depot_data[depot]
        _cell(ws, row, 1, depot, h_align="left", bg=alt_fill, border=True)
        row_total = 0.0
        has_error = False
        error_msgs = []
        for ci, product in enumerate(PRODUCTS, 2):
            info    = prod_info.get(product, {})
            balance = info.get("balance")
            error   = info.get("error")
            if balance is not None:
                _cell(ws, row, ci, balance, num_fmt="#,##0", bg=alt_fill, border=True)
                row_total += balance
                product_totals[product] += balance
            else:
                c = ws.cell(row, ci, "N/A")
                c.font = _font(bold=False, color="FF0000")
                c.alignment = _align()
                c.border = _border()
                if alt_fill:
                    c.fill = _fill(alt_fill)
                has_error = True
                if error:
                    error_msgs.append(f"{product}: {error}")
        total_col = len(PRODUCTS) + 2
        if not has_error or row_total > 0:
            _cell(ws, row, total_col, row_total, bold=True, num_fmt="#,##0", bg=alt_fill, border=True)
        else:
            c = ws.cell(row, total_col, "N/A")
            c.font = _font(bold=False, color="FF0000")
            c.alignment = _align()
            c.border = _border()
        status_col = len(PRODUCTS) + 3
        if not has_error:
            status_txt, status_clr = "✓ OK", "00AA00"
        elif row_total > 0:
            status_txt, status_clr = f"⚠ Partial ({'; '.join(error_msgs)})", "B8860B"
        else:
            status_txt, status_clr = f"✗ {'; '.join(error_msgs) or 'No data'}", "CC0000"
        c = ws.cell(row, status_col, status_txt)
        c.font = _font(bold=False, color=status_clr, size=10)
        c.alignment = _align(h="left")
        c.border = _border()
        if alt_fill:
            c.fill = _fill(alt_fill)
        row += 1
    row += 1
    ws.cell(row, 1, "PRODUCT TOTALS").font = _font(bold=True, color=HEADER_FG, size=11)
    ws.cell(row, 1).fill = _fill(DARK_BLUE)
    ws.cell(row, 1).alignment = _align(h="left")
    ws.cell(row, 1).border = _border()
    grand_total = 0.0
    for ci, product in enumerate(PRODUCTS, 2):
        total = product_totals[product]
        grand_total += total
        c = ws.cell(row, ci, total)
        c.font = _font(bold=True, color=HEADER_FG, size=11)
        c.fill = _fill(DARK_BLUE)
        c.number_format = "#,##0"
        c.alignment = _align()
        c.border = _border()
    total_col = len(PRODUCTS) + 2
    c = ws.cell(row, total_col, grand_total)
    c.font = _font(bold=True, color=HEADER_FG, size=12)
    c.fill = _fill(DARK_BLUE)
    c.number_format = "#,##0"
    c.alignment = _align()
    c.border = _border()
    ws.cell(row, len(PRODUCTS) + 3, "").fill = _fill(DARK_BLUE)
    ws.cell(row, len(PRODUCTS) + 3, "").border = _border()
    ws.column_dimensions["A"].width = 42
    for ci, _ in enumerate(PRODUCTS, 2):
        ws.column_dimensions[get_column_letter(ci)].width = 18
    ws.column_dimensions[get_column_letter(total_col)].width = 22
    ws.column_dimensions[get_column_letter(len(PRODUCTS) + 3)].width = 38


def write_daily_orders_sheet(ws, daily_df: pd.DataFrame):
    """Write the OilCorp daily orders to an Excel worksheet."""
    if daily_df.empty:
        ws.cell(1, 1, "No daily order data available.")
        return

    ws.cell(1, 1, "OILCORP ENERGIA LIMITED — DAILY ORDERS").font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(daily_df.columns))

    for ci, col in enumerate(daily_df.columns, 1):
        c = ws.cell(2, ci, col)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(MED_BLUE)
        c.alignment = _align()
        c.border = _border()

    for ri, row_data in daily_df.iterrows():
        excel_row = ri + 3
        alt_fill  = "F2F2F2" if ri % 2 == 0 else None
        for ci, val in enumerate(row_data, 1):
            c = ws.cell(excel_row, ci, val)
            c.border = _border()
            c.alignment = _align(h="left" if ci in [1, 3, 5, 6] else "center")
            if alt_fill:
                c.fill = _fill(alt_fill)
            if isinstance(val, (int, float)) and ci in [7, 8]:
                c.number_format = "#,##0.00"

    for ci, col in enumerate(daily_df.columns, 1):
        ws.column_dimensions[get_column_letter(ci)].width = max(14, len(str(col)) + 4)


def generate_excel(tables, summary_df, month_label,
                   opening_data=None, closing_data=None, daily_df=None):
    wb = Workbook()
    ws_p1 = wb.active
    ws_p1.title = "P1"
    write_p1_sheet(ws_p1, tables)
    ws_sum = wb.create_sheet("SUMMARY")
    write_summary_sheet(ws_sum, summary_df)
    if opening_data:
        ws_open = wb.create_sheet("OPENING STOCK")
        write_stock_balance_sheet(ws_open, opening_data, "OPENING", month_label)
    if closing_data:
        ws_close = wb.create_sheet("CLOSING STOCK")
        write_stock_balance_sheet(ws_close, closing_data, "CLOSING", month_label)
    if daily_df is not None and not daily_df.empty:
        ws_daily = wb.create_sheet("DAILY ORDERS")
        write_daily_orders_sheet(ws_daily, daily_df)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════
# DEFAULT DEPOTS
# ══════════════════════════════════════════════════════════════
DEFAULT_DEPOTS = [
    "BOST GLOBAL DEPOT",
    "CHASE PETROLEUM - TEMA",
    "PETROLEUM HUB LIMITED",
    "PETROLEUM WARE HOUSE AND SUPPLIES LIMITED",
    "PLATON OIL GAS GHANA LIMITED",
    "QUANTUM LPG LOGISTICS LIMITED",
    "QUANTUM OIL TERMINAL LIMITED",
    "SENTUO OIL REFINERY - TEMA",
    "TEMA FUEL COMPANY (TFC)",
    "TEMA MULTI PRODUCTS (TMPT)",
    "TEMA OIL REFINERY (TOR)",
    "TEMA OIL TERMINAL PLC",
    "VANA ENERGY LIMITED TEMA",
]

def _init_depot_list():
    if "configured_depots" not in st.session_state:
        st.session_state["configured_depots"] = list(DEFAULT_DEPOTS)


# ══════════════════════════════════════════════════════════════
# PAGE: ORDER REQUEST REPORT
# ══════════════════════════════════════════════════════════════
def page_order_report():
    st.markdown("""
    <div class="hero-header">
      <div class="hero-inner">
        <div class="hero-eyebrow">▸ OILCORP ENERGIA LIMITED</div>
        <h1 class="hero-title">Order Request <span>Report</span></h1>
        <p class="hero-subtitle">Upload an ORDER REQUEST sheet to generate P1, Summary and Stock Balance reports.</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── File Upload ───────────────────────────────────────────
    st.markdown('<div class="section-label">01 — INPUT FILE</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Drop your Excel file here or click to browse",
        type=["xlsx"], label_visibility="visible",
        key="order_upload",
    )

    if not uploaded:
        st.markdown("""
        <div style="text-align:center; padding: 3rem 0; color: var(--text-muted);">
            <div style="font-size:2.5rem; margin-bottom:0.8rem;">📂</div>
            <div style="font-family:'JetBrains Mono',monospace; font-size:0.75rem; letter-spacing:0.1em; text-transform:uppercase;">
                Awaiting .xlsx upload
            </div>
        </div>
        """, unsafe_allow_html=True)
        return

    with st.spinner("Parsing workbook…"):
        try:
            df = load_order_data(uploaded)
        except Exception as e:
            st.error(f"Could not read ORDER REQUEST sheet: {e}")
            return

    df["Year"]  = df["DATE"].dt.year
    df["Month"] = df["DATE"].dt.month
    years = sorted(df["Year"].unique(), reverse=True)

    # ── Period Selector ───────────────────────────────────────
    st.markdown('<div class="section-label">02 — REPORTING PERIOD</div>', unsafe_allow_html=True)
    with st.container():
        st.markdown('<div class="period-card">', unsafe_allow_html=True)
        col1, col2, _ = st.columns([1, 1, 2])
        with col1:
            sel_year = st.selectbox("Year", years, key="or_year")
        with col2:
            avail_months = sorted(df[df["Year"] == sel_year]["Month"].unique())
            sel_month = st.selectbox("Month", avail_months, format_func=lambda m: MONTHS[m], key="or_month")
        st.markdown('</div>', unsafe_allow_html=True)

    month_label = f"{MONTHS[sel_month]} {sel_year}"
    filtered    = df[(df["Year"] == sel_year) & (df["Month"] == sel_month)].copy()

    if filtered.empty:
        st.warning("No data found for the selected period.")
        return

    w1 = filtered[filtered["DATE"].dt.day <= 15]
    w2 = filtered[filtered["DATE"].dt.day >= 16]
    total_qty = filtered["Quantity"].sum()
    n_omcs    = filtered["OMC"].nunique()
    n_depots  = filtered["Depot"].nunique()

    st.markdown(f"""
    <div class="stat-grid">
      <div class="stat-card blue">
        <div class="stat-label">Total Records</div>
        <div class="stat-value">{len(filtered):,}</div>
        <div class="stat-sub">W1: {len(w1):,} &nbsp;·&nbsp; W2: {len(w2):,}</div>
      </div>
      <div class="stat-card gold">
        <div class="stat-label">Total Volume (L)</div>
        <div class="stat-value">{total_qty:,.0f}</div>
        <div class="stat-sub">{month_label}</div>
      </div>
      <div class="stat-card green">
        <div class="stat-label">OMCs · Depots</div>
        <div class="stat-value">{n_omcs} <span style="font-size:1.2rem;color:var(--text-muted)">·</span> {n_depots}</div>
        <div class="stat-sub">Active this period</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    tables     = build_p1_tables(filtered)
    summary_df = build_summary(filtered)

    st.markdown('<div class="section-label">03 — REPORT VIEWS</div>', unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs([
        "  P1 — Depot Breakdown  ",
        "  Loading Summary  ",
        "  Stock Balances  ",
        "  Export  ",
    ])

    # ── TAB 1: P1 ─────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="tab-heading">P1 — Depot / Product / Window</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="tab-subhead">Order requests by depot, product, and fortnightly window — {month_label}</div>', unsafe_allow_html=True)
        if not tables:
            st.info("No data to display for this period.")
        else:
            depots_in_data = sorted(set(t["depot"] for t in tables))
            for depot in depots_in_data:
                with st.expander(f"🏭  {depot}", expanded=True):
                    depot_tables = [t for t in tables if t["depot"] == depot]
                    cols = st.columns(min(len(depot_tables), 2))
                    for i, tbl in enumerate(depot_tables):
                        with cols[i % 2]:
                            icon = {"PMS":"🟡","AGO":"🔵","LPG":"🟢"}.get(tbl["product"],"⚪")
                            st.markdown(f"**{icon} {tbl['title']}**")
                            display = tbl["data"][["OMC", "Quantity"]].copy()
                            display.columns = ["OMC", "Quantity (L)"]
                            display["Quantity (L)"] = display["Quantity (L)"].apply(lambda x: f"{x:,.0f}")
                            total = tbl["data"]["Quantity"].sum()
                            total_row = pd.DataFrame([{"OMC":"GRAND TOTAL","Quantity (L)":f"{total:,.0f}"}])
                            display = pd.concat([display, total_row], ignore_index=True)
                            st.dataframe(display, use_container_width=True, hide_index=True)

    # ── TAB 2: SUMMARY ────────────────────────────────────────
    with tab2:
        st.markdown('<div class="tab-heading">Loading Summary</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="tab-subhead">Aggregate volumes by depot group and product — {month_label}</div>', unsafe_allow_html=True)
        display_sum = summary_df.copy()
        for col in ["AGO", "PMS", "LPG", "GRAND TOTAL"]:
            if col in display_sum.columns:
                display_sum[col] = display_sum[col].apply(
                    lambda x: f"{int(x):,}" if pd.notna(x) and x != 0 else "—"
                )
        col_rename = {"DepotGroup": "DEPOT"} if "DepotGroup" in display_sum.columns else {}
        st.dataframe(display_sum.rename(columns=col_rename), use_container_width=True, hide_index=True)

    # ── TAB 3: STOCK BALANCES ─────────────────────────────────
    with tab3:
        _init_depot_list()
        st.markdown('<div class="tab-heading">OILCORP Stock Balances</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="tab-subhead">NPA stock transaction ledger for OILCORP ENERGIA LIMITED — {month_label}</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="info-panel">
            <strong>Opening Balance</strong> — First <em>Stock Take</em> after the b/fwd row (if present), otherwise the <em>Balance b/fwd</em> value.<br>
            <strong>Closing Balance</strong> — Last recorded running balance at end of the selected month.
        </div>
        """, unsafe_allow_html=True)

        col_manager, col_controls = st.columns([5, 4], gap="large")

        with col_manager:
            configured   = st.session_state["configured_depots"]
            n_configured = len(configured)
            st.markdown(f"""
            <div class="depot-manager">
              <div class="depot-manager-header">
                <span class="depot-manager-title">Depot List</span>
                <span class="depot-count-badge">{n_configured} depots</span>
              </div>
              <div class="depot-list-body">
            """ + "".join([
                f'<div class="depot-row"><span class="depot-row-index">{i+1}</span>'
                f'<span class="depot-check">✓</span><span class="depot-name">{d}</span></div>'
                for i, d in enumerate(configured)
            ]) + "</div></div>", unsafe_allow_html=True)

            with st.expander("✏️  Edit depot list", expanded=False):
                st.markdown('<div class="add-depot-title">Remove a depot</div>', unsafe_allow_html=True)
                if configured:
                    remove_choice = st.selectbox("Select depot to remove", options=configured,
                                                  key="depot_remove_choice", label_visibility="collapsed")
                    if st.button("🗑  Remove selected depot", key="btn_remove_depot"):
                        st.session_state["configured_depots"].remove(remove_choice)
                        st.rerun()
                else:
                    st.caption("No depots configured.")

                st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
                st.markdown('<div class="add-depot-title">Add a depot</div>', unsafe_allow_html=True)
                addable = [d for d in sorted(DEPOT_MAP.keys()) if d not in st.session_state["configured_depots"]]
                if addable:
                    add_choice = st.selectbox("Choose from .env depots", options=addable,
                                               key="depot_add_choice", label_visibility="collapsed")
                    if st.button("➕  Add depot", key="btn_add_depot"):
                        st.session_state["configured_depots"].append(add_choice)
                        st.session_state["configured_depots"].sort()
                        st.rerun()
                else:
                    st.caption("All .env depots are already in the list.")

                st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
                st.markdown('<div class="add-depot-title">Custom depot name</div>', unsafe_allow_html=True)
                custom_name = st.text_input("Type a depot name manually", placeholder="e.g. GHANA OIL CO.LTD, TAKORADI",
                                             key="depot_custom_input", label_visibility="collapsed")
                if st.button("➕  Add custom depot", key="btn_add_custom") and custom_name.strip():
                    name = custom_name.strip()
                    if name not in st.session_state["configured_depots"]:
                        st.session_state["configured_depots"].append(name)
                        st.session_state["configured_depots"].sort()
                        st.rerun()
                    else:
                        st.warning("That depot is already in the list.")

                st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
                if st.button("↺  Reset to defaults", key="btn_reset_depots"):
                    st.session_state["configured_depots"] = list(DEFAULT_DEPOTS)
                    st.rerun()

        with col_controls:
            selected_products = st.multiselect("Products", ["PMS", "GASOIL", "LPG"],
                                                default=["PMS", "GASOIL", "LPG"], key="oilcorp_products")
            depots_to_query   = st.session_state["configured_depots"]
            products_to_query = selected_products if selected_products else ["PMS", "GASOIL", "LPG"]
            total_calls       = len(depots_to_query) * len(products_to_query)

            st.markdown(f"""
            <div class="query-counter">
                <div class="query-number">{total_calls}</div>
                <div>API calls &nbsp;·&nbsp; <strong>{len(depots_to_query)}</strong> depot(s) × <strong>{len(products_to_query)}</strong> product(s) for <strong>{month_label}</strong></div>
            </div>
            """, unsafe_allow_html=True)

            if not depots_to_query:
                st.warning("Add at least one depot to the list before fetching.")
            elif st.button("🔄  Fetch Stock Balances", type="primary", use_container_width=True):
                opening_results, closing_results = [], []
                progress  = st.progress(0, text="Initialising…")
                log_box   = st.empty()
                log_lines = []
                call_n    = 0

                for depot_name in depots_to_query:
                    depot_id = DEPOT_MAP.get(depot_name)
                    if not depot_id:
                        for prod in products_to_query:
                            opening_results.append({"depot": depot_name, "product": prod, "balance": None, "error": "Depot ID not found in .env"})
                            closing_results.append({"depot": depot_name, "product": prod, "balance": None, "error": "Depot ID not found in .env"})
                        call_n += len(products_to_query)
                        log_lines.append(f"⚠️ {depot_name} — not in .env, skipped")
                        log_box.markdown(f"<div class='log-console'>{'<br>'.join(log_lines[-12:])}</div>", unsafe_allow_html=True)
                        continue

                    for product in products_to_query:
                        product_id = PRODUCT_MAP.get(product)
                        call_n    += 1
                        progress.progress(call_n / total_calls, text=f"Fetching {depot_name} · {product}  ({call_n}/{total_calls})")
                        result = fetch_oilcorp_stock_balances(sel_year, sel_month, depot_name, depot_id, product, product_id)
                        if result["opening"] is not None:
                            log_lines.append(f"✅ {depot_name} [{product}] — Open: {result['opening']:,.0f} | Close: {result['closing']:,.0f}")
                        else:
                            log_lines.append(f"⚠️ {depot_name} [{product}] — {result['error']}")
                        log_box.markdown(f"<div class='log-console'>{'<br>'.join(log_lines[-12:])}</div>", unsafe_allow_html=True)
                        opening_results.append({"depot": depot_name, "product": product, "balance": result["opening"], "error": result["error"]})
                        closing_results.append({"depot": depot_name, "product": product, "balance": result["closing"], "error": result["error"]})

                progress.progress(1.0, text="✅ Complete")
                st.session_state["oilcorp_opening"] = opening_results
                st.session_state["oilcorp_closing"] = closing_results
                n_ok = sum(1 for r in opening_results if r["balance"] is not None)
                st.success(f"Fetched {total_calls} combinations — **{n_ok}** with data, **{total_calls - n_ok}** with errors.")

        # Display results
        opening_data = st.session_state.get("oilcorp_opening")
        closing_data = st.session_state.get("oilcorp_closing")

        if opening_data or closing_data:
            PRODUCTS = ["PMS", "GASOIL", "LPG"]

            def _pivot_for_display(data):
                depot_order, depot_map_local = [], {}
                for entry in data:
                    d, p = entry["depot"], entry["product"]
                    if d not in depot_map_local:
                        depot_map_local[d] = {}
                        depot_order.append(d)
                    depot_map_local[d][p] = entry
                rows = []
                for depot in depot_order:
                    prod_info = depot_map_local[depot]
                    row_dict  = {"DEPOT": depot}
                    row_total = 0.0
                    has_error = False
                    for prod in PRODUCTS:
                        info    = prod_info.get(prod, {})
                        balance = info.get("balance")
                        error   = info.get("error", "")
                        if balance is not None:
                            row_dict[prod] = f"{balance:,.0f}"
                            row_total     += balance
                        else:
                            row_dict[prod] = "N/A"
                            has_error = True
                    row_dict["TOTAL"]  = f"{row_total:,.0f}" if (not has_error or row_total > 0) else "N/A"
                    row_dict["STATUS"] = "✓" if not has_error else "⚠ Partial" if row_total > 0 else "✗"
                    rows.append(row_dict)
                totals = {p: 0.0 for p in PRODUCTS}
                for entry in data:
                    if entry["balance"] is not None:
                        totals[entry["product"]] = totals.get(entry["product"], 0.0) + entry["balance"]
                total_row = {"DEPOT": "GRAND TOTAL"}
                grand = 0.0
                for prod in PRODUCTS:
                    total_row[prod] = f"{totals[prod]:,.0f}"
                    grand += totals[prod]
                total_row["TOTAL"]  = f"{grand:,.0f}"
                total_row["STATUS"] = ""
                rows.append(total_row)
                return pd.DataFrame(rows)

            st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
            col_a, col_b = st.columns(2)
            with col_a:
                if opening_data:
                    st.markdown("#### 📂 Opening Stock Balance")
                    st.dataframe(_pivot_for_display(opening_data), use_container_width=True, hide_index=True)
            with col_b:
                if closing_data:
                    st.markdown("#### 📁 Closing Stock Balance")
                    st.dataframe(_pivot_for_display(closing_data), use_container_width=True, hide_index=True)

    # ── TAB 4: EXPORT ─────────────────────────────────────────
    with tab4:
        opening_data = st.session_state.get("oilcorp_opening")
        closing_data = st.session_state.get("oilcorp_closing")
        daily_df_exp = st.session_state.get("oilcorp_daily_df", pd.DataFrame())
        has_stock    = bool(opening_data or closing_data)
        has_daily    = isinstance(daily_df_exp, pd.DataFrame) and not daily_df_exp.empty

        p1_badge     = '<span class="sheet-badge active">P1</span>'
        sum_badge    = '<span class="sheet-badge active">SUMMARY</span>'
        open_badge   = f'<span class="sheet-badge {"active" if opening_data else ""}">OPENING STOCK</span>'
        close_badge  = f'<span class="sheet-badge {"active" if closing_data else ""}">CLOSING STOCK</span>'
        daily_badge  = f'<span class="sheet-badge {"active" if has_daily else ""}">DAILY ORDERS</span>'

        st.markdown(f"""
        <div class="export-card">
            <span class="export-icon">📊</span>
            <div class="export-title">Excel Workbook</div>
            <div class="export-desc">
                Generate a formatted .xlsx report with all active sheets for <strong>{month_label}</strong>.
            </div>
            <div class="sheet-badges">
                {p1_badge} {sum_badge} {open_badge} {close_badge} {daily_badge}
            </div>
        </div>
        """, unsafe_allow_html=True)

        if not has_stock:
            st.markdown("""
            <div class="info-panel" style="max-width:520px; margin: 0 auto 1rem auto; text-align:center;">
                Visit the <strong>Stock Balances</strong> tab to fetch OILCORP stock data.<br>
                Visit <strong>Daily Orders</strong> in the sidebar to include order data in the export.
            </div>
            """, unsafe_allow_html=True)

        col_left, col_mid, col_right = st.columns([1, 2, 1])
        with col_mid:
            if st.button("⚡  Build Excel Report", type="primary", use_container_width=True):
                with st.spinner("Compiling workbook…"):
                    excel_buf = generate_excel(
                        tables, summary_df, month_label,
                        opening_data=opening_data,
                        closing_data=closing_data,
                        daily_df=daily_df_exp if has_daily else None,
                    )
                fname = f"OilCorp_Report_{MONTHS[sel_month]}_{sel_year}.xlsx"
                st.download_button(
                    label=f"⬇  Download  {fname}",
                    data=excel_buf,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# ══════════════════════════════════════════════════════════════
# PAGE: DAILY ORDERS — OilCorp Energia
# ══════════════════════════════════════════════════════════════
def page_daily_orders():
    st.markdown("""
    <div class="hero-header">
      <div class="hero-inner">
        <div class="hero-eyebrow">▸ OILCORP ENERGIA LIMITED</div>
        <h1 class="hero-title">Daily <span>Orders</span></h1>
        <p class="hero-subtitle">Truck-level dispatch records for OilCorp Energia — fetched live from the NPA portal.</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-label">01 — DATE RANGE</div>', unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="period-card">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            start_date = st.date_input("From", value=datetime.now() - timedelta(days=7), key="daily_start")
        with col2:
            end_date   = st.date_input("To",   value=datetime.now(),                     key="daily_end")
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            st.info(f"**{(end_date - start_date).days + 1}** day window selected")
        st.markdown('</div>', unsafe_allow_html=True)

    if start_date > end_date:
        st.error("Start date must be before end date.")
        return

    st.markdown('<div class="section-label">02 — FETCH</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="info-panel">
        Queries the NPA Daily Order Report endpoint using <strong>OilCorp Energia's</strong> user credentials
        and filters results to only show orders where OilCorp is the listed BDC.
        Covers all products (PMS, GASOIL, LPG) and all depots simultaneously.
    </div>
    """, unsafe_allow_html=True)

    col_btn, _ = st.columns([1, 3])
    with col_btn:
        fetch_clicked = st.button("🔄  Fetch Daily Orders", type="primary", use_container_width=True)

    if fetch_clicked:
        with st.spinner(f"Fetching OilCorp daily orders ({start_date.strftime('%d %b')} → {end_date.strftime('%d %b %Y')})…"):
            try:
                df_fetched = fetch_oilcorp_daily_orders(
                    datetime.combine(start_date, datetime.min.time()),
                    datetime.combine(end_date,   datetime.min.time()),
                )
                st.session_state["oilcorp_daily_df"]    = df_fetched
                st.session_state["oilcorp_daily_start"] = start_date
                st.session_state["oilcorp_daily_end"]   = end_date
                if df_fetched.empty:
                    st.warning("No OilCorp daily order records found for this date range. "
                               "Check credentials and date range, or the NPA portal may be temporarily unavailable.")
                else:
                    st.success(f"✅ {len(df_fetched):,} order records retrieved for OilCorp Energia.")
            except Exception as e:
                st.error(f"Fetch failed: {e}")
                return

    df = st.session_state.get("oilcorp_daily_df", pd.DataFrame())
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        st.markdown("""
        <div style="text-align:center; padding: 3rem 0; color: var(--text-muted);">
            <div style="font-size:2.5rem; margin-bottom:0.8rem;">📋</div>
            <div style="font-family:'JetBrains Mono',monospace; font-size:0.75rem; letter-spacing:0.1em; text-transform:uppercase;">
                Select a date range and fetch to see orders
            </div>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── Metrics ────────────────────────────────────────────────
    st.markdown('<div class="section-label">03 — SUMMARY</div>', unsafe_allow_html=True)

    d_start = st.session_state.get("oilcorp_daily_start", start_date)
    d_end   = st.session_state.get("oilcorp_daily_end",   end_date)
    period_lbl = f"{d_start.strftime('%d %b')} – {d_end.strftime('%d %b %Y')}"

    total_vol = df["Quantity (L)"].sum() if "Quantity (L)" in df.columns else 0
    n_orders  = len(df)
    n_depots  = df["Depot"].nunique() if "Depot" in df.columns else 0
    total_val = (df["Quantity (L)"] * df["Price (₵/L)"]).sum() if all(c in df.columns for c in ["Quantity (L)", "Price (₵/L)"]) else 0

    st.markdown(f"""
    <div class="daily-metric-row">
      <div class="daily-metric">
        <div class="daily-metric-label">Total Orders</div>
        <div class="daily-metric-value">{n_orders:,}</div>
        <div class="daily-metric-sub">{period_lbl}</div>
      </div>
      <div class="daily-metric">
        <div class="daily-metric-label">Total Volume</div>
        <div class="daily-metric-value">{total_vol:,.0f}</div>
        <div class="daily-metric-sub">Litres dispatched</div>
      </div>
      <div class="daily-metric">
        <div class="daily-metric-label">Depots Active</div>
        <div class="daily-metric-value">{n_depots}</div>
        <div class="daily-metric-sub">Unique depots</div>
      </div>
      <div class="daily-metric">
        <div class="daily-metric-label">Est. Value</div>
        <div class="daily-metric-value">₵{total_val:,.0f}</div>
        <div class="daily-metric-sub">Volume × price</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Product breakdown ──────────────────────────────────────
    st.markdown('<div class="section-label">04 — BREAKDOWN</div>', unsafe_allow_html=True)

    tab_prod, tab_depot, tab_omc, tab_detail = st.tabs([
        "  By Product  ", "  By Depot  ", "  By OMC  ", "  Full Detail  "
    ])

    with tab_prod:
        st.markdown('<div class="tab-heading">Product Summary</div>', unsafe_allow_html=True)
        if "Product" in df.columns and "Quantity (L)" in df.columns:
            prod_grp = (
                df.groupby("Product")
                .agg(Orders=("Order Number", "count"), Volume=("Quantity (L)", "sum"))
                .reset_index()
                .sort_values("Volume", ascending=False)
            )
            prod_grp["Volume"] = prod_grp["Volume"].apply(lambda x: f"{x:,.0f}")
            st.dataframe(prod_grp, use_container_width=True, hide_index=True)
        else:
            st.info("Product column not available.")

    with tab_depot:
        st.markdown('<div class="tab-heading">Depot Summary</div>', unsafe_allow_html=True)
        if "Depot" in df.columns and "Quantity (L)" in df.columns:
            depot_grp = (
                df.groupby("Depot")
                .agg(Orders=("Order Number","count"), Volume=("Quantity (L)","sum"))
                .reset_index()
                .sort_values("Volume", ascending=False)
            )
            depot_grp["Volume"] = depot_grp["Volume"].apply(lambda x: f"{x:,.0f}")
            st.dataframe(depot_grp, use_container_width=True, hide_index=True)
        else:
            st.info("Depot column not available.")

    with tab_omc:
        st.markdown('<div class="tab-heading">OMC Summary</div>', unsafe_allow_html=True)
        if "OMC" in df.columns and "Quantity (L)" in df.columns:
            omc_grp = (
                df.groupby("OMC")
                .agg(Orders=("Order Number","count"), Volume=("Quantity (L)","sum"))
                .reset_index()
                .sort_values("Volume", ascending=False)
                .rename(columns={"Volume":"Total Volume (L)"})
            )
            omc_grp["Total Volume (L)"] = omc_grp["Total Volume (L)"].apply(lambda x: f"{x:,.0f}")
            st.dataframe(omc_grp, use_container_width=True, hide_index=True)
        else:
            st.info("OMC column not available.")

    with tab_detail:
        st.markdown('<div class="tab-heading">Full Order Detail</div>', unsafe_allow_html=True)
        # Filter controls
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            products_all = ["All"] + sorted(df["Product"].dropna().unique().tolist()) if "Product" in df.columns else ["All"]
            prod_filter  = st.selectbox("Filter by Product", products_all, key="daily_prod_filter")
        with fc2:
            depots_all  = ["All"] + sorted(df["Depot"].dropna().unique().tolist()) if "Depot" in df.columns else ["All"]
            depot_filter = st.selectbox("Filter by Depot", depots_all, key="daily_depot_filter")
        with fc3:
            status_all  = ["All"] + sorted(df["Status"].dropna().unique().tolist()) if "Status" in df.columns else ["All"]
            status_filter = st.selectbox("Filter by Status", status_all, key="daily_status_filter")

        filt = df.copy()
        if prod_filter   != "All" and "Product" in filt.columns: filt = filt[filt["Product"]  == prod_filter]
        if depot_filter  != "All" and "Depot" in filt.columns:   filt = filt[filt["Depot"]    == depot_filter]
        if status_filter != "All" and "Status" in filt.columns:  filt = filt[filt["Status"]   == status_filter]

        filt_vol = filt["Quantity (L)"].sum() if "Quantity (L)" in filt.columns else 0
        st.caption(f"Showing **{len(filt):,}** records  |  Volume: **{filt_vol:,.0f} L**")
        st.dataframe(filt, use_container_width=True, hide_index=True, height=450)

    # ── Download ───────────────────────────────────────────────
    st.markdown('<div class="section-label">05 — EXPORT</div>', unsafe_allow_html=True)
    col_l, col_m, col_r = st.columns([1, 2, 1])
    with col_m:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Daily Orders", index=False)
            if "Product" in df.columns:
                for prod in df["Product"].unique():
                    prod_df = df[df["Product"] == prod]
                    if not prod_df.empty:
                        prod_df.to_excel(writer, sheet_name=prod[:31], index=False)
        buf.seek(0)
        fname = f"OilCorp_DailyOrders_{d_start.strftime('%Y%m%d')}_{d_end.strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label=f"⬇  Download {fname}",
            data=buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ══════════════════════════════════════════════════════════════
# SIDEBAR NAV + MAIN
# ══════════════════════════════════════════════════════════════
def main():
    with st.sidebar:
        st.markdown("""
        <div style="padding: 1.5rem 0 1rem 0; text-align: center;">
            <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.6rem;
                        letter-spacing: 0.25em; text-transform: uppercase; color: #2F81F7;
                        margin-bottom: 0.5rem;">▸ OILCORP ENERGIA</div>
            <div style="font-family: 'Barlow Condensed', sans-serif; font-size: 1.5rem;
                        font-weight: 800; color: #E6EDF3; letter-spacing: -0.01em;
                        line-height: 1;">Intelligence<br><span style="color:#2F81F7">Suite</span></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<hr style="border-color: #30363D; margin: 0.5rem 0 1.2rem 0;">', unsafe_allow_html=True)

        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.6rem;
                    letter-spacing: 0.2em; text-transform: uppercase; color: #484F58;
                    margin-bottom: 0.6rem;">Navigation</div>
        """, unsafe_allow_html=True)

        page = st.radio(
            "nav",
            options=["📋  Order Request Report", "📦  Daily Orders"],
            label_visibility="collapsed",
            key="nav_page",
        )

        # ── Data status badges ─────────────────────────────────
        st.markdown('<hr style="border-color: #30363D; margin: 1.5rem 0 1rem 0;">', unsafe_allow_html=True)
        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.6rem;
                    letter-spacing: 0.2em; text-transform: uppercase; color: #484F58;
                    margin-bottom: 0.8rem;">Session Data</div>
        """, unsafe_allow_html=True)

        has_opening = bool(st.session_state.get("oilcorp_opening"))
        has_closing = bool(st.session_state.get("oilcorp_closing"))
        has_daily   = isinstance(st.session_state.get("oilcorp_daily_df"), pd.DataFrame) and not st.session_state.get("oilcorp_daily_df", pd.DataFrame()).empty

        def _badge(label, active):
            color = "#3FB950" if active else "#484F58"
            icon  = "●" if active else "○"
            return f'<div style="font-size: 0.8rem; color: {color}; margin-bottom: 0.3rem;">{icon} {label}</div>'

        st.markdown(
            _badge("Opening Stock", has_opening) +
            _badge("Closing Stock", has_closing) +
            _badge("Daily Orders",  has_daily),
            unsafe_allow_html=True,
        )

        if any([has_opening, has_closing, has_daily]):
            st.markdown('<br>', unsafe_allow_html=True)
            if st.button("🗑  Clear session data", use_container_width=True):
                for k in ["oilcorp_opening", "oilcorp_closing", "oilcorp_daily_df",
                           "oilcorp_daily_start", "oilcorp_daily_end"]:
                    st.session_state.pop(k, None)
                st.rerun()

        st.markdown('<hr style="border-color: #30363D; margin: 1.5rem 0 1rem 0;">', unsafe_allow_html=True)
        st.markdown("""
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.6rem;
                    letter-spacing: 0.1em; color: #484F58; text-align: center;">
            OILCORP ENERGIA LIMITED<br>NPA Reporting Suite
        </div>
        """, unsafe_allow_html=True)

    if page == "📋  Order Request Report":
        page_order_report()
    elif page == "📦  Daily Orders":
        page_daily_orders()


main()