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

st.set_page_config(page_title="Order Request Report", layout="wide")

# ── Styling helpers ────────────────────────────────────────────────────────────
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

# ── Data loading ───────────────────────────────────────────────────────────────
def load_order_data(file):
    df = pd.read_excel(file, sheet_name="ORDER REQUEST", header=8)
    df = df[["DATE", "Name of OMC", "Product", "Depot", "Quantity", "Comments"]].copy()
    df.columns = ["DATE", "OMC", "Product", "Depot", "Quantity", "Comments"]
    df = df.dropna(subset=["DATE", "OMC", "Depot", "Quantity"])
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df = df.dropna(subset=["DATE"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    return df

# ── P1-style pivot builder ─────────────────────────────────────────────────────
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
                    "depot": depot,
                    "product": product,
                    "window": window,
                    "data": pivot,
                })
    return tables

# ── Summary builder ────────────────────────────────────────────────────────────
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
# OILCORP STOCK BALANCE — ENV + CONFIG
# ══════════════════════════════════════════════════════════════

OILCORP_USER_ID = os.getenv("OILCORP_USER_ID", "123293")
OILCORP_BDC_ID  = os.getenv("OILCORP_BDC_ID",  "20900")
NPA_COMPANY_ID  = os.getenv("NPA_COMPANY_ID",   "1")

STOCK_TXN_URL = os.getenv(
    "NPA_STOCK_TRANSACTION_URL",
    "https://iml.npa-enterprise.com/NewNPA/home/CreateStockTransactionReport"
)

PRODUCT_MAP = {
    "PMS":    int(os.getenv("PRODUCT_PREMIUM_ID", "12")),
    "GASOIL": int(os.getenv("PRODUCT_GASOIL_ID",  "14")),
    "LPG":    int(os.getenv("PRODUCT_LPG_ID",     "28")),
}

# All depots from .env — load dynamically
def _load_depot_map() -> dict:
    """
    Read all DEPOT_* keys from .env and return {display_name: id}.
    Handles the naming quirks in the .env file.
    """
    depot_map = {}
    for key, value in os.environ.items():
        if not key.startswith("DEPOT_"):
            continue
        raw = key[6:]  # strip "DEPOT_"
        # Convert underscores to spaces then apply known fixes
        name = raw.replace("_", " ").strip()
        # Apply the same normalisation as the original npa_dashboard.py
        if name == "GHANA OIL COLTD TAKORADI":
            name = "GHANA OIL CO.LTD, TAKORADI"
        elif name == "GOIL LPG BOTTLING PLANT TEMA":
            name = "GOIL LPG BOTTLING PLANT -TEMA"
        elif name == "GOIL LPG BOTTLING PLANT KUMASI":
            name = "GOIL LPG BOTTLING PLANT- KUMASI"
        elif name == "NEWGAS CYLINDER BOTTLING LIMITED TEMA":
            name = "NEWGAS CYLINDER BOTTLING LIMITED-TEMA"
        elif name == "CHASE PETROLEUM TEMA":
            name = "CHASE PETROLEUM - TEMA"
        elif name == "TEMA FUEL COMPANY TFC":
            name = "TEMA FUEL COMPANY (TFC)"
        elif name == "TEMA MULTI PRODUCTS TMPT":
            name = "TEMA MULTI PRODUCTS (TMPT)"
        elif name == "TEMA OIL REFINERY TOR":
            name = "TEMA OIL REFINERY (TOR)"
        elif name == "GHANA OIL COMPANY LTD SEKONDI NAVAL BASE":
            name = "GHANA OIL COMPANY LTD (SEKONDI NAVAL BASE)"
        elif name == "GHANSTOCK LIMITED TAKORADI":
            name = "GHANSTOCK LIMITED (TAKORADI)"
        elif name == "SENTUO OIL REFINERY TEMA":
            name = "SENTUO OIL REFINERY - TEMA"
        elif name == "TAKORADI BLUE OCEAN INVESTMENT LIMITED":
            name = "TAKORADI BLUE OCEAN INVESTMENT LIMITED"
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


# Matches any line that starts with DD/MM/YYYY followed by whitespace
_DATE_LINE_RE = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(.*)")

# Matches the two trailing numbers (volume + balance) at the end of a ledger line.
# Balance may be negative e.g. (8,453)
_TAIL_RE = re.compile(
    r"(\([\d,]+\)|[\d,]+)\s+(\([\d,]+\)|[\d,]+)\s*$"
)


def _parse_any_date_line(line: str) -> dict | None:
    """
    Parse ANY dated ledger line, including 'Balance b/fwd'.

    Format:  DD/MM/YYYY  TRANS#  DESCRIPTION  [ACCOUNT]  VOLUME  BALANCE

    The tricky case is 'Balance b/fwd': the regex captures 'Balance' as the
    trans token and 'b/fwd 93,547 93,547' as rest.  We therefore do a
    two-pass search:
      Pass 1 — try matching descriptions against `rest`  (normal rows)
      Pass 2 — try matching against `trans + ' ' + rest` (catches b/fwd)
    The first pass that yields a known description wins.
    """
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
    """
    Parse all ledger lines from the PDF, INCLUDING 'Balance b/fwd'.
    Returns list sorted by appearance order (page → line).
    """
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
# OPENING / CLOSING STOCK FETCHER FOR OILCORP
# ══════════════════════════════════════════════════════════════

def fetch_oilcorp_stock_balances(
    year: int, month: int, depot_name: str, depot_id: int, product: str, product_id: int
) -> dict:
    """
    Fetch the stock transaction PDF for OILCORP ENERGIA for the given month,
    depot and product.  Returns:
        {
          "opening": float | None,   # Opening balance for this month
          "closing": float | None,   # Last running balance recorded in the ledger
          "records": list,
          "error":   str | None,
        }

    Opening balance rules
    ---------------------
    1. If a "Stock Take" row exists anywhere after the "Balance b/fwd" row,
       use the Balance of the FIRST "Stock Take" row as the opening balance.
    2. Otherwise, fall back to the Balance column of the "Balance b/fwd" row.

    This handles the common case where a stock take is performed at the start
    of the month and supersedes the carried-forward balance.

    Closing balance rules
    ---------------------
    Walk the records in reverse and take the Balance of the last row that is
    NOT a b/fwd row.  If the only record present is the b/fwd itself (e.g. no
    activity during the month), we fall back to that same balance — the account
    opened and closed at the same value.
    """
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

    # ── Opening: prefer first Stock Take after the b/fwd; fall back to b/fwd ──
    bfwd_balance     = None
    bfwd_index       = None
    first_stock_take = None

    for i, r in enumerate(records):
        if r["Description"] == "Balance b/fwd" and bfwd_balance is None:
            bfwd_balance = float(r["Balance"])
            bfwd_index   = i
        elif (
            r["Description"] == "Stock Take"
            and bfwd_index is not None          # must come AFTER the b/fwd row
            and i > bfwd_index
            and first_stock_take is None
        ):
            first_stock_take = float(r["Balance"])

    # Use Stock Take balance if one was found, otherwise use b/fwd balance
    opening = first_stock_take if first_stock_take is not None else bfwd_balance

    # ── Closing: last non-b/fwd row; fallback to last row if none ────────────
    closing = None
    for r in reversed(records):
        if r["Description"] != "Balance b/fwd":
            closing = float(r["Balance"])
            break

    # Edge case: only the b/fwd row exists (no activity during the month)
    # → opening and closing are the same value
    if closing is None and records:
        closing = float(records[-1]["Balance"])

    return {"opening": opening, "closing": closing, "records": records, "error": None}


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
            c.font      = _font(bold=True, color=HEADER_FG)
            c.fill      = _fill(MED_BLUE)
            c.alignment = _align()
            c.border    = _border()
        row += 1

        start_data = row
        for _, r in data.iterrows():
            ws.cell(row, 1, r["OMC"]).alignment = _align(h="left")
            ws.cell(row, 1).border = _border()
            for ci, val in enumerate([r["Quantity"], r["TOTAL"]], 2):
                c = ws.cell(row, ci, val)
                c.number_format = "#,##0"
                c.alignment     = _align()
                c.border        = _border()
            row += 1

        gt_cell = ws.cell(row, 1, "GRAND TOTAL")
        gt_cell.font      = _font(bold=True)
        gt_cell.fill      = _fill(YELLOW)
        gt_cell.border    = _border()
        gt_cell.alignment = _align(h="left")
        for ci in [2, 3]:
            col_letter = get_column_letter(ci)
            c = ws.cell(row, ci, f"=SUM({col_letter}{start_data}:{col_letter}{row-1})")
            c.font          = _font(bold=True)
            c.fill          = _fill(YELLOW)
            c.number_format = "#,##0"
            c.alignment     = _align()
            c.border        = _border()
        row += 3

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18


def write_summary_sheet(ws, summary_df):
    ws.cell(1, 1, "LOADING SUMMARY").font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill      = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)

    headers = ["DEPOT", "AGO", "PMS", "LPG", "GRAND TOTAL"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(2, ci, h)
        c.font      = _font(bold=True, color=HEADER_FG)
        c.fill      = _fill(MED_BLUE)
        c.alignment = _align()
        c.border    = _border()

    for ri, row_data in summary_df.iterrows():
        excel_row = ri + 3
        is_total  = str(row_data.iloc[0]) == "GRAND TOTAL"
        fill      = _fill(ORANGE) if is_total else _fill(GREEN) if ri % 2 == 0 else PatternFill()
        for ci, val in enumerate(row_data, 1):
            c = ws.cell(excel_row, ci, val)
            c.border    = _border()
            c.alignment = _align(h="left" if ci == 1 else "center")
            c.fill      = fill
            if ci > 1:
                c.number_format = "#,##0"
            if is_total:
                c.font = _font(bold=True)

    for col, width in zip(["A", "B", "C", "D", "E"], [20, 14, 14, 10, 16]):
        ws.column_dimensions[col].width = width


def _cell(ws, row, col, value, bold=False, bg=None, fg="000000",
          size=11, h_align="center", num_fmt=None, border=True):
    c = ws.cell(row, col, value)
    c.font      = _font(bold=bold, color=fg, size=size)
    c.alignment = _align(h=h_align)
    if bg:
        c.fill = _fill(bg)
    if border:
        c.border = _border()
    if num_fmt:
        c.number_format = num_fmt
    return c


def write_stock_balance_sheet(ws, balance_data: list, sheet_type: str, month_label: str):
    """
    Write either Opening Stock or Closing Stock sheet.
    balance_data = list of dicts:
      { depot, product, balance, error }
    sheet_type = "OPENING" | "CLOSING"
    """
    # ── Title ──
    title_text = f"OILCORP ENERGIA LIMITED — {sheet_type} STOCK BALANCE ({month_label})"
    ws.cell(1, 1, title_text).font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill      = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)

    subtitle = (
        "Balance carried forward at start of month (Stock Take if available, else b/fwd)"
        if sheet_type == "OPENING"
        else "Last recorded running balance at end of month"
    )
    ws.cell(2, 1, subtitle).font      = _font(bold=False, color="595959", size=10)
    ws.cell(2, 1).fill                = _fill(LIGHT_BLUE)
    ws.cell(2, 1).alignment           = _align()
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)

    # ── Column headers ──
    headers = ["DEPOT", "PRODUCT", "BALANCE (LT / KG)", "STATUS"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(3, ci, h)
        c.font      = _font(bold=True, color=HEADER_FG, size=11)
        c.fill      = _fill(MED_BLUE)
        c.alignment = _align()
        c.border    = _border()

    # ── Data rows ──
    row = 4
    product_totals: dict[str, float] = {}
    depot_totals:   dict[str, float] = {}

    for entry in balance_data:
        depot   = entry["depot"]
        product = entry["product"]
        balance = entry["balance"]
        error   = entry["error"]

        alt_fill = "F2F2F2" if row % 2 == 0 else None

        _cell(ws, row, 1, depot,   h_align="left",   bg=alt_fill, border=True)
        _cell(ws, row, 2, product, h_align="center",  bg=alt_fill, border=True)

        if balance is not None:
            _cell(ws, row, 3, balance, num_fmt="#,##0", bg=alt_fill, border=True)
            product_totals[product] = product_totals.get(product, 0) + balance
            depot_totals[depot]     = depot_totals.get(depot, 0)     + balance
        else:
            c = ws.cell(row, 3, "N/A")
            c.font      = _font(bold=False, color="FF0000")
            c.alignment = _align()
            c.border    = _border()
            if alt_fill:
                c.fill = _fill(alt_fill)

        status_txt = "✓ OK" if balance is not None else f"✗ {error or 'No data'}"
        status_clr = "00AA00" if balance is not None else "CC0000"
        c = ws.cell(row, 4, status_txt)
        c.font      = _font(bold=False, color=status_clr, size=10)
        c.alignment = _align(h="center")
        c.border    = _border()
        if alt_fill:
            c.fill = _fill(alt_fill)

        row += 1

    # ── Product subtotals block ──
    row += 1
    ws.cell(row, 1, "PRODUCT TOTALS").font      = _font(bold=True, color=HEADER_FG, size=12)
    ws.cell(row, 1).fill                         = _fill("1F3864")
    ws.cell(row, 1).alignment                    = _align(h="left")
    ws.cell(row, 1).border                       = _border()
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1

    prod_order = ["PMS", "GASOIL", "LPG"]
    for prod in prod_order:
        total = product_totals.get(prod, 0)
        _cell(ws, row, 1, f"Total {prod}", bold=True, h_align="left",
              bg=LIGHT_BLUE, border=True)
        ws.cell(row, 1).font = _font(bold=True, color="1F3864", size=11)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        _cell(ws, row, 3, total, bold=True, num_fmt="#,##0", bg=LIGHT_BLUE, border=True)
        ws.cell(row, 3).font = _font(bold=True, color="1F3864", size=11)
        _cell(ws, row, 4, "LT / KG", bold=False, bg=LIGHT_BLUE, border=True)
        row += 1

    # Grand total
    grand = sum(product_totals.values())
    _cell(ws, row, 1, "GRAND TOTAL", bold=True, fg=HEADER_FG, bg=DARK_BLUE,
          h_align="left", border=True, size=12)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    _cell(ws, row, 3, grand, bold=True, fg=HEADER_FG, bg=DARK_BLUE,
          num_fmt="#,##0", border=True, size=12)
    _cell(ws, row, 4, "LT / KG", bold=True, fg=HEADER_FG, bg=DARK_BLUE,
          border=True, size=12)

    # ── Column widths ──
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22


def generate_excel(tables, summary_df, month_label,
                   opening_data=None, closing_data=None):
    wb = Workbook()

    # P1
    ws_p1 = wb.active
    ws_p1.title = "P1"
    write_p1_sheet(ws_p1, tables)

    # SUMMARY
    ws_sum = wb.create_sheet("SUMMARY")
    write_summary_sheet(ws_sum, summary_df)

    # OPENING STOCK (optional)
    if opening_data:
        ws_open = wb.create_sheet("OPENING STOCK")
        write_stock_balance_sheet(ws_open, opening_data, "OPENING", month_label)

    # CLOSING STOCK (optional)
    if closing_data:
        ws_close = wb.create_sheet("CLOSING STOCK")
        write_stock_balance_sheet(ws_close, closing_data, "CLOSING", month_label)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════

st.title("📦 Order Request Report Generator")
st.markdown(
    "Upload an Excel file with an **ORDER REQUEST** sheet to generate "
    "P1, Summary, and — optionally — Opening & Closing Stock Balance reports."
)

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded:
    with st.spinner("Reading file…"):
        try:
            df = load_order_data(uploaded)
        except Exception as e:
            st.error(f"Could not read ORDER REQUEST sheet: {e}")
            st.stop()

    df["Year"]  = df["DATE"].dt.year
    df["Month"] = df["DATE"].dt.month

    years = sorted(df["Year"].unique(), reverse=True)
    months = {
        1:"January", 2:"February", 3:"March",    4:"April",
        5:"May",     6:"June",     7:"July",      8:"August",
        9:"September",10:"October",11:"November",12:"December"
    }

    col1, col2 = st.columns(2)
    with col1:
        sel_year = st.selectbox("Select Year", years)
    with col2:
        avail_months = sorted(df[df["Year"] == sel_year]["Month"].unique())
        sel_month    = st.selectbox("Select Month", avail_months,
                                    format_func=lambda m: months[m])

    month_label = f"{months[sel_month]} {sel_year}"

    filtered = df[(df["Year"] == sel_year) & (df["Month"] == sel_month)].copy()

    if filtered.empty:
        st.warning("No data found for the selected period.")
        st.stop()

    w1 = filtered[filtered["DATE"].dt.day <= 15]
    w2 = filtered[filtered["DATE"].dt.day >= 16]

    st.success(
        f"**{month_label}** — {len(filtered):,} records | "
        f"W1: {len(w1):,} | W2: {len(w2):,}"
    )

    tables     = build_p1_tables(filtered)
    summary_df = build_summary(filtered)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 P1 — Depot Breakdown",
        "📊 Summary",
        "🏦 OILCORP Stock Balances",
        "⬇️ Export to Excel",
    ])

    with tab1:
        st.subheader(f"P1 — Depot / Product / Window Breakdown ({month_label})")
        if not tables:
            st.info("No data to display.")
        else:
            depots_in_data = sorted(set(t["depot"] for t in tables))
            for depot in depots_in_data:
                with st.expander(f"🏭 {depot}", expanded=True):
                    depot_tables = [t for t in tables if t["depot"] == depot]
                    cols = st.columns(min(len(depot_tables), 2))
                    for i, tbl in enumerate(depot_tables):
                        with cols[i % 2]:
                            st.markdown(f"**{tbl['title']}**")
                            display = tbl["data"][["OMC", "Quantity"]].copy()
                            display.columns = ["OMC", "Quantity (L)"]
                            display["Quantity (L)"] = display["Quantity (L)"].apply(
                                lambda x: f"{x:,.0f}"
                            )
                            total = tbl["data"]["Quantity"].sum()
                            total_row = pd.DataFrame(
                                [{"OMC": "GRAND TOTAL", "Quantity (L)": f"{total:,.0f}"}]
                            )
                            display = pd.concat([display, total_row], ignore_index=True)
                            st.dataframe(display, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader(f"Loading Summary — {month_label}")
        display_sum = summary_df.copy()
        for col in ["AGO", "PMS", "LPG", "GRAND TOTAL"]:
            if col in display_sum.columns:
                display_sum[col] = display_sum[col].apply(
                    lambda x: f"{int(x):,}" if pd.notna(x) and x != 0 else "-"
                )
        col_rename = {"DepotGroup": "DEPOT"} if "DepotGroup" in display_sum.columns else {}
        st.dataframe(display_sum.rename(columns=col_rename),
                     use_container_width=True, hide_index=True)

    # ── Tab 3 — OILCORP STOCK BALANCES ───────────────────────────────────────
    with tab3:
        st.subheader(f"🏦 OILCORP ENERGIA — Opening & Closing Stock Balance ({month_label})")

        st.markdown("""
        Fetches the NPA stock transaction ledger for **OILCORP ENERGIA LIMITED** only —
        for every combination of depot and product that is configured in your `.env` file.

        - **Opening Balance** = first *Stock Take* value after the b/fwd row (if present), otherwise the *Balance b/fwd* value.
        - **Closing Balance** = the last running balance entry at the end of the selected month.
        """)

        # Let user pick which depots & products to query
        available_depots = sorted(DEPOT_MAP.keys())
        if not available_depots:
            st.error(
                "No depots found in your `.env` file. "
                "Please ensure DEPOT_* keys are set correctly."
            )
            st.stop()

        selected_depots = st.multiselect(
            "Select depots to query (leave blank to query ALL configured depots)",
            available_depots,
            default=[],
            key="oilcorp_depots",
        )
        selected_products = st.multiselect(
            "Select products",
            ["PMS", "GASOIL", "LPG"],
            default=["PMS", "GASOIL", "LPG"],
            key="oilcorp_products",
        )

        depots_to_query   = selected_depots if selected_depots else available_depots
        products_to_query = selected_products if selected_products else ["PMS", "GASOIL", "LPG"]

        st.info(
            f"Will query **{len(depots_to_query)} depot(s)** × "
            f"**{len(products_to_query)} product(s)** = "
            f"**{len(depots_to_query) * len(products_to_query)} API calls** "
            f"for {month_label}."
        )

        if st.button("🔄 Fetch OILCORP Stock Balances", type="primary"):
            opening_results: list[dict] = []
            closing_results: list[dict] = []
            errors: list[str] = []

            total_calls = len(depots_to_query) * len(products_to_query)
            progress    = st.progress(0, text="Starting…")
            log_box     = st.empty()
            log_lines: list[str] = []
            call_n = 0

            for depot_name in depots_to_query:
                depot_id = DEPOT_MAP.get(depot_name)
                if not depot_id:
                    for prod in products_to_query:
                        opening_results.append({"depot": depot_name, "product": prod,
                                                "balance": None, "error": "Depot ID not found"})
                        closing_results.append({"depot": depot_name, "product": prod,
                                                "balance": None, "error": "Depot ID not found"})
                    call_n += len(products_to_query)
                    continue

                for product in products_to_query:
                    product_id = PRODUCT_MAP.get(product)
                    call_n += 1
                    progress.progress(
                        call_n / total_calls,
                        text=f"Fetching {depot_name} — {product} ({call_n}/{total_calls})…"
                    )

                    result = fetch_oilcorp_stock_balances(
                        sel_year, sel_month, depot_name, depot_id, product, product_id
                    )

                    if result["opening"] is not None:
                        log_lines.append(
                            f"✅ {depot_name} [{product}] — Open: {result['opening']:,.0f}"
                        )
                    else:
                        log_lines.append(
                            f"⚠️ {depot_name} [{product}] — {result['error']}"
                        )

                    # ── FIX: dark background + light text so log is always readable ──
                    log_box.markdown(
                        "<div style='"
                        "font-family:monospace;"
                        "font-size:12px;"
                        "background:#1e1e1e;"
                        "color:#d4d4d4;"
                        "padding:10px 14px;"
                        "border-radius:6px;"
                        "max-height:180px;"
                        "overflow-y:auto;"
                        "line-height:1.6;"
                        "'>"
                        + "<br>".join(log_lines[-10:])
                        + "</div>",
                        unsafe_allow_html=True,
                    )

                    opening_results.append({
                        "depot":   depot_name,
                        "product": product,
                        "balance": result["opening"],
                        "error":   result["error"],
                    })
                    closing_results.append({
                        "depot":   depot_name,
                        "product": product,
                        "balance": result["closing"],
                        "error":   result["error"],
                    })

            progress.progress(1.0, text="✅ Done")

            st.session_state["oilcorp_opening"] = opening_results
            st.session_state["oilcorp_closing"] = closing_results

            n_ok = sum(1 for r in opening_results if r["balance"] is not None)
            st.success(
                f"Fetched {total_calls} combinations — "
                f"**{n_ok}** with data, "
                f"**{total_calls - n_ok}** no data / error."
            )

        # ── Display results if available ──────────────────────────────────────
        opening_data = st.session_state.get("oilcorp_opening")
        closing_data = st.session_state.get("oilcorp_closing")

        if opening_data or closing_data:
            col_a, col_b = st.columns(2)

            def _render_balance_table(data: list, label: str):
                rows = []
                for entry in data:
                    bal = entry["balance"]
                    rows.append({
                        "Depot":       entry["depot"],
                        "Product":     entry["product"],
                        "Balance (LT)": f"{bal:,.0f}" if bal is not None else "—",
                        "Status":      "✓" if bal is not None else f"✗ {entry['error']}",
                    })
                df_display = pd.DataFrame(rows)
                st.markdown(f"#### {label}")
                st.dataframe(df_display, use_container_width=True, hide_index=True)

                # Product totals
                totals = {}
                for entry in data:
                    if entry["balance"] is not None:
                        totals[entry["product"]] = totals.get(entry["product"], 0) + entry["balance"]
                if totals:
                    st.markdown("**Product Totals:**")
                    tot_df = pd.DataFrame([
                        {"Product": p, "Total (LT/KG)": f"{v:,.0f}"}
                        for p, v in sorted(totals.items())
                    ])
                    tot_df.loc[len(tot_df)] = {
                        "Product": "GRAND TOTAL",
                        "Total (LT/KG)": f"{sum(totals.values()):,.0f}"
                    }
                    st.dataframe(tot_df, use_container_width=True, hide_index=True)

            with col_a:
                _render_balance_table(opening_data, "📂 Opening Stock Balance")
            with col_b:
                _render_balance_table(closing_data, "📁 Closing Stock Balance")

    # ── Tab 4 — EXPORT ────────────────────────────────────────────────────────
    with tab4:
        st.subheader("Export Report to Excel")
        st.markdown(
            "Generates an Excel workbook with **P1**, **SUMMARY**, and — if fetched — "
            "**OPENING STOCK** and **CLOSING STOCK** sheets for OILCORP ENERGIA."
        )

        opening_data = st.session_state.get("oilcorp_opening")
        closing_data = st.session_state.get("oilcorp_closing")

        if opening_data or closing_data:
            st.success(
                "✅ OILCORP stock balance data is ready — "
                "it will be included in the exported workbook."
            )
        else:
            st.info(
                "ℹ️ Go to the **OILCORP Stock Balances** tab first to fetch "
                "opening/closing data before exporting."
            )

        if st.button("📥 Generate Excel Report", type="primary"):
            with st.spinner("Building Excel file…"):
                excel_buf = generate_excel(
                    tables, summary_df, month_label,
                    opening_data=opening_data,
                    closing_data=closing_data,
                )
            fname = f"Report_{months[sel_month]}_{sel_year}.xlsx"
            st.download_button(
                label=f"⬇️ Download {fname}",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )