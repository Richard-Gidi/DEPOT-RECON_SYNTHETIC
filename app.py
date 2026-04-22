import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    """Return list of dicts: {title, pivot_df} for each Depot/Product/Window."""
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
                # For BOST depots keep sub-depot info (BOST-APD → APD etc.)
                # Single depot tables: just OMC × quantity
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
    """Collapse sub-depots for BOST, then build Depot × Product totals."""
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

# ── Excel writer ───────────────────────────────────────────────────────────────
def write_p1_sheet(ws, tables):
    row = 1
    for tbl in tables:
        data = tbl["data"]
        title = tbl["title"]

        # Title row
        ws.cell(row, 1, title).font = _font(bold=True, color=HEADER_FG, size=12)
        ws.cell(row, 1).fill = _fill(DARK_BLUE)
        ws.cell(row, 1).alignment = _align()
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        row += 1

        # Header
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

        # Grand total
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
        row += 3  # gap

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18


def write_summary_sheet(ws, summary_df):
    ws.cell(1, 1, "LOADING SUMMARY").font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)

    # Sub-header row (product labels)
    headers = ["DEPOT", "AGO", "PMS", "LPG", "GRAND TOTAL"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(2, ci, h)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(MED_BLUE)
        c.alignment = _align()
        c.border = _border()

    # Data
    for ri, row_data in summary_df.iterrows():
        excel_row = ri + 3
        is_total = str(row_data.iloc[0]) == "GRAND TOTAL"
        fill = _fill(ORANGE) if is_total else _fill(GREEN) if ri % 2 == 0 else PatternFill()
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


# ── Opening/Closing stock balance sheets ───────────────────────────────────────
import re
import threading
import calendar as _cal
import concurrent.futures
import requests as _requests
import pdfplumber

DEPOT_MAP = {
    "ADINKRA STORAGE COMPANY GHANA LIMITED":         241,
    "AKWAABA LINK INVESTMENTS LIMITED":              20538,
    "BLUE OCEAN CYLINDER BOTTLING PLANT":            20937,
    "BLUE OCEAN INVESTMENT LTD KOTOKA AIRPORT ATK":  20507,
    "BOST - ACCRA PLAINS":                           20458,
    "BOST - AKOSOMBO":                               20463,
    "BOST - BOLGATANGA":                             20461,
    "BOST - BUIPE":                                  20460,
    "BOST - KUMASI":                                 20459,
    "BOST - MAMIWATER":                              20462,
    "BOST GLOBAL DEPOT":                             20901,
    "BULK OIL STORAGE AND TRANSPORTATION COMPANY":   243,
    "CHASE PETROLEUM - TEMA":                        141,
    "GHANA BUNKERING SERVICES":                      20615,
    "GHANA NATIONAL GAS COMPANY LIMITED":            20465,
    "GHANA OIL CO.LTD, TAKORADI":                    239,
    "GHANA OIL COMPANY LTD (SEKONDI NAVAL BASE)":    20492,
    "GHANSTOCK LIMITED (TAKORADI)":                  20510,
    "GOIL LPG BOTTLING PLANT -TEMA":                 20887,
    "GOIL LPG BOTTLING PLANT- KUMASI":               20888,
    "MATRIX GAS GHANA LIMITED":                      20852,
    "NEWGAS CYLINDER BOTTLING LIMITED-TEMA":         20922,
    "OLD BAUXITE JETTY":                             20450,
    "PETROLEUM HUB LIMITED":                         20774,
    "PETROLEUM WARE HOUSE AND SUPPLIES":             142,
    "PLATON OIL AND GAS":                            20464,
    "QUANTUM LPG LOGISTICS LIMITED":                 20850,
    "QUANTUM OIL TERMINAL LIMITED":                  20639,
    "QUANTUM TERMINALS LIMITED":                     238,
    "RIDGE ENERGY LIMITED":                          20485,
    "SENTUO OIL REFINERY- TEMA":                     20918,
    "TAKORADI BLUE OCEAN INVESTMENT LIMITED":        20467,
    "TEMA FUEL COMPANY (TFC)":                       145,
    "TEMA MULTI PRODUCTS (TMPT)":                    20477,
    "TEMA OIL REFINERY (TOR)":                       237,
    "TEMA OIL TERMINAL PLC":                         20838,
    "TOTAL PETROLEUM GHANA LIMITED":                 240,
    "VANA ENERGY LIMITED TEMA":                      366,
    "ZEN TERMINALS LIMITED":                         143,
}

PRODUCT_MAP = {
    "PMS": 12,
    "AGO": 14,
    "LPG": 28,
}

BDC_ENTITY_MAP = {
    "GOIL COMPANY LIMITED":             20900,
    "TOTAL PETROLEUM GHANA LIMITED":    20901,
    "SHELL GHANA LIMITED":              20902,
    "PUMA ENERGY GHANA PLC":            20903,
    "ENGEN PETROLEUM GHANA LIMITED":    20904,
    "STAR OIL GHANA LIMITED":           20905,
    "FIRST LIGHT ENERGY LIMITED":       20906,
    "ORYX ENERGIES GHANA LIMITED":      20907,
    "METRO PETROLEUM LIMITED":          20908,
    "SENTUO OIL REFINERY":              20918,
}

NPA_USER_ID   = "123292"
STOCK_TXN_URL = "https://iml.npa-enterprise.com/NewNPA/home/CreateStockTransactionReport"

_HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/pdf,text/html,*/*;q=0.8",
}

_DESCRIPTIONS_SORTED = sorted([
    "Balance b/fwd", "Stock Take", "Sale",
    "Custody Transfer In", "Custody Transfer Out", "Product Outturn",
], key=len, reverse=True)

_SKIP_PFX = (
    "national petroleum authority", "stock transaction report",
    "bdc :", "depot :", "product :", "printed by", "printed on",
    "date trans #", "actual stock balance", "stock commitments",
    "available stock balance", "last stock update", "i.t.s from",
)


def _fetch_pdf_bytes(url, params, timeout=90):
    try:
        r = _requests.get(url, params=params, headers=_HTTP_HEADERS, timeout=timeout)
        r.raise_for_status()
        return r.content if r.content[:4] == b"%PDF" else None
    except Exception:
        return None


def _parse_stock_transaction_pdf(pdf_bytes):
    def _skip(line):
        lo = line.strip().lower()
        return lo.startswith(_SKIP_PFX) or bool(re.match(r"^\d{1,2}\s+\w+,\s+\d{4}", line.strip()))

    def _pnum(s):
        s = s.strip()
        neg = s.startswith("(") and s.endswith(")")
        try:
            v = int(s.strip("()").replace(",", ""))
            return -v if neg else v
        except ValueError:
            return None

    def _parse_line(line):
        line = line.strip()
        if not re.match(r"^\d{2}/\d{2}/\d{4}\b", line):
            return None
        parts = line.split()
        date, trans = parts[0], (parts[1] if len(parts) > 1 else "")
        rest = line[len(date):].strip()[len(trans):].strip()
        desc = after = None
        for d in _DESCRIPTIONS_SORTED:
            if rest.lower().startswith(d.lower()):
                desc, after = d, rest[len(d):].strip()
                break
        if desc is None or desc == "Balance b/fwd":
            return None
        nums = re.findall(r"\([\d,]+\)|[\d,]+", after)
        if len(nums) < 2:
            return None
        vol, bal = _pnum(nums[-2]), _pnum(nums[-1])
        trail = re.search(re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"\s*$", after)
        acct = after[:trail.start()].strip() if trail else " ".join(after.split()[:-2])
        return {"Date": date, "Trans #": trans, "Description": desc,
                "Account": acct, "Volume": vol or 0, "Balance": bal or 0}

    records = []
    seen = set()

    # ── CHANGE 1: also parse Balance b/fwd lines (needed for opening logic) ──
    def _parse_line_full(line):
        line = line.strip()
        if not re.match(r"^\d{2}/\d{2}/\d{4}\b", line):
            return None
        parts = line.split()
        date, trans = parts[0], (parts[1] if len(parts) > 1 else "")
        rest = line[len(date):].strip()[len(trans):].strip()
        desc = after = None
        for d in _DESCRIPTIONS_SORTED:
            if rest.lower().startswith(d.lower()):
                desc, after = d, rest[len(d):].strip()
                break
        if desc is None:
            return None
        nums = re.findall(r"\([\d,]+\)|[\d,]+", after)
        if len(nums) < 2:
            return None
        vol, bal = _pnum(nums[-2]), _pnum(nums[-1])
        trail = re.search(re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"\s*$", after)
        acct = after[:trail.start()].strip() if trail else " ".join(after.split()[:-2])
        return {"Date": date, "Trans #": trans, "Description": desc,
                "Account": acct, "Volume": vol or 0, "Balance": bal or 0}

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
                    row = _parse_line_full(line)
                    if row:
                        key = (row["Date"], row["Trans #"], row["Description"], row["Volume"])
                        if key not in seen:
                            seen.add(key)
                            records.append(row)
    except Exception:
        pass
    return records


def _get_opening_closing(records):
    """
    Opening balance logic (CHANGE 1):
      - Find the first 'Balance b/fwd'.
      - If the very next record is a 'Stock Take', use that Stock Take's balance
        (it supersedes the b/fwd figure).
      - Otherwise use the Balance b/fwd balance directly.
      - If no Balance b/fwd exists, fall back to the first record's balance.

    Closing balance:
      Balance from the last transaction record.
    """
    if not records:
        return {"opening": 0, "closing": 0}

    opening = None
    for i, rec in enumerate(records):
        if rec["Description"] == "Balance b/fwd":
            if i + 1 < len(records) and records[i + 1]["Description"] == "Stock Take":
                opening = records[i + 1]["Balance"]
            else:
                opening = rec["Balance"]
            break

    if opening is None:
        opening = records[0]["Balance"]

    closing = records[-1]["Balance"]
    return {"opening": opening, "closing": closing}


def write_stock_balance_sheet(ws, df, title, balance_col, month_label):
    num_cols = 4
    ws.cell(1, 1, title).font = _font(bold=True, color=HEADER_FG, size=14)
    ws.cell(1, 1).fill = _fill(DARK_BLUE)
    ws.cell(1, 1).alignment = _align()
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)

    ws.cell(2, 1, f"Period: {month_label}").font = _font(bold=False, color=HEADER_FG, size=11)
    ws.cell(2, 1).fill = _fill(MED_BLUE)
    ws.cell(2, 1).alignment = _align(h="left")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)

    for ci, h in enumerate(["BDC", "DEPOT", "PRODUCT", "BALANCE (LT/KG)"], 1):
        c = ws.cell(3, ci, h)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(MED_BLUE)
        c.alignment = _align()
        c.border = _border()

    if df.empty:
        ws.cell(4, 1, "No data available — check BDC/Depot/Product combinations and API connectivity.")
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=num_cols)
        for col, width in zip(["A", "B", "C", "D"], [40, 40, 12, 20]):
            ws.column_dimensions[col].width = width
        return

    row = 4
    alternating = False
    for _, r in df.iterrows():
        fill_row = _fill(LIGHT_BLUE) if alternating else PatternFill()
        for ci, val in enumerate(
            [r.get("BDC", ""), r.get("Depot", ""), r.get("Product", ""), r.get(balance_col, 0)], 1
        ):
            c = ws.cell(row, ci, val)
            c.border = _border()
            c.fill = fill_row
            if ci < 4:
                c.alignment = _align(h="left")
                c.font = _font(size=10)
            else:
                c.alignment = _align()
                c.number_format = "#,##0"
                c.font = _font(size=10)
        alternating = not alternating
        row += 1

    row += 1
    ws.cell(row, 1, "PRODUCT TOTALS").font = _font(bold=True, color=HEADER_FG)
    ws.cell(row, 1).fill = _fill(DARK_BLUE)
    ws.cell(row, 1).alignment = _align()
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    row += 1

    for ci, h in enumerate(["PRODUCT", "TOTAL BALANCE (LT/KG)", "BDC COUNT", "DEPOT COUNT"], 1):
        c = ws.cell(row, ci, h)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(MED_BLUE)
        c.alignment = _align()
        c.border = _border()
    row += 1

    grand_grand = 0.0
    for product in sorted(df["Product"].unique()):
        sub = df[df["Product"] == product]
        total_bal = float(sub[balance_col].sum())
        grand_grand += total_bal
        for ci, val in enumerate(
            [product, total_bal, sub["BDC"].nunique(), sub["Depot"].nunique()], 1
        ):
            c = ws.cell(row, ci, val)
            c.border = _border()
            c.font = _font(bold=(ci == 1))
            c.alignment = _align(h="left" if ci == 1 else "center")
            if ci == 2:
                c.number_format = "#,##0"
        row += 1

    for ci in range(1, 5):
        c = ws.cell(row, ci)
        c.fill = _fill(ORANGE)
        c.border = _border()
        c.font = _font(bold=True)
    ws.cell(row, 1, "GRAND TOTAL").alignment = _align(h="left")
    gt_val = ws.cell(row, 2, grand_grand)
    gt_val.number_format = "#,##0"
    gt_val.alignment = _align()

    for col, width in zip(["A", "B", "C", "D"], [38, 42, 12, 22]):
        ws.column_dimensions[col].width = width


def generate_excel(tables, summary_df, month_label, opening_df=None, closing_df=None):
    wb = Workbook()
    ws_p1 = wb.active
    ws_p1.title = "P1"
    write_p1_sheet(ws_p1, tables)

    ws_sum = wb.create_sheet("SUMMARY")
    write_summary_sheet(ws_sum, summary_df)

    ws_open = wb.create_sheet("OPENING STOCK BALANCE")
    write_stock_balance_sheet(
        ws_open,
        opening_df if opening_df is not None else pd.DataFrame(),
        "OPENING STOCK BALANCE", "Opening (LT)", month_label,
    )

    ws_close = wb.create_sheet("CLOSING STOCK BALANCE")
    write_stock_balance_sheet(
        ws_close,
        closing_df if closing_df is not None else pd.DataFrame(),
        "CLOSING STOCK BALANCE", "Closing (LT)", month_label,
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ───────────────────────────────────────────────────────────────
st.title("📦 Order Request Report Generator")
st.markdown("Upload an Excel file with an **ORDER REQUEST** sheet to generate P1 & Summary reports.")

# ── CHANGE 2: Fix white-on-white progress bar ──────────────────────────────────
st.markdown("""
<style>
/* Force progress bar text to be dark and visible */
[data-testid="stProgressBar"] p,
[data-testid="stProgressBar"] span,
[data-testid="stProgressBar"] label {
    color: #1a1a1a !important;
    font-weight: 600 !important;
}
/* Give the progress container a light grey background */
[data-testid="stProgressBar"] {
    background-color: #eef0f5 !important;
    border-radius: 6px;
    padding: 4px 10px;
}
/* Dark blue filled bar */
[data-testid="stProgressBar"] > div > div {
    background-color: #1F3864 !important;
}
/* Fetch log box — dark background so white text is readable */
.fetch-log-box {
    background: #1e2130;
    border: 1px solid #3a4060;
    border-radius: 8px;
    padding: 10px 14px;
    font-family: monospace;
    font-size: 12px;
    color: #e0e0e0 !important;
    max-height: 180px;
    overflow-y: auto;
    margin-top: 6px;
}
</style>
""", unsafe_allow_html=True)

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
        1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
        7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"
    }

    col1, col2 = st.columns(2)
    with col1:
        sel_year = st.selectbox("Select Year", years)
    with col2:
        avail_months = sorted(df[df["Year"] == sel_year]["Month"].unique())
        sel_month = st.selectbox("Select Month", avail_months, format_func=lambda m: months[m])

    month_label = f"{months[sel_month]} {sel_year}"

    filtered = df[(df["Year"] == sel_year) & (df["Month"] == sel_month)].copy()

    if filtered.empty:
        st.warning("No data found for the selected period.")
        st.stop()

    w1 = filtered[filtered["DATE"].dt.day <= 15]
    w2 = filtered[filtered["DATE"].dt.day >= 16]

    st.success(f"**{month_label}** — {len(filtered):,} records | W1: {len(w1):,} | W2: {len(w2):,}")

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["📋 P1 — Depot Breakdown", "📊 Summary", "⬇️ Export to Excel"])

    tables = build_p1_tables(filtered)
    summary_df = build_summary(filtered)

    with tab1:
        st.subheader(f"P1 — Depot / Product / Window Breakdown ({month_label})")
        if not tables:
            st.info("No data to display.")
        else:
            # Group by depot for cleaner display
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
                            display["Quantity (L)"] = display["Quantity (L)"].apply(lambda x: f"{x:,.0f}")
                            total = tbl["data"]["Quantity"].sum()
                            total_row = pd.DataFrame([{"OMC": "GRAND TOTAL", "Quantity (L)": f"{total:,.0f}"}])
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
        st.dataframe(display_sum.rename(columns=col_rename), use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("Export Report to Excel")

        # ── Stock balance fetch section ────────────────────────────────────────
        st.markdown("#### Step 1 — Fetch Opening & Closing Stock Balances (optional)")
        st.markdown("Queries the NPA Stock Transaction API for each BDC × Depot × Product over the full calendar month.")

        api_user_id = st.text_input(
            "NPA API User ID (lngUserId)",
            value=NPA_USER_ID,
            help="The user ID passed to the NPA Stock Transaction API.",
        )

        with st.expander("⚙️ Advanced: Limit fetch scope (optional)", expanded=False):
            sel_bdcs     = st.multiselect("BDCs (blank = all)",     list(BDC_ENTITY_MAP.keys()), key="exp_bdcs")
            sel_depots   = st.multiselect("Depots (blank = all)",   list(DEPOT_MAP.keys()),      key="exp_depots")
            sel_products = st.multiselect("Products (blank = all)", list(PRODUCT_MAP.keys()),    key="exp_prods")

        fetch_bdcs     = sel_bdcs     or list(BDC_ENTITY_MAP.keys())
        fetch_depots   = sel_depots   or list(DEPOT_MAP.keys())
        fetch_products = sel_products or list(PRODUCT_MAP.keys())

        total_combos = len(fetch_bdcs) * len(fetch_depots) * len(fetch_products)
        st.caption(f"Will query **{total_combos:,}** BDC × Depot × Product combinations.")

        if st.button("📡 Fetch Stock Balances", key="fetch_stock_bal"):
            last_day  = _cal.monthrange(sel_year, sel_month)[1]
            start_str = f"{sel_month:02d}/01/{sel_year}"
            end_str   = f"{sel_month:02d}/{last_day:02d}/{sel_year}"

            tasks = [
                (bdc_n, BDC_ENTITY_MAP[bdc_n], dep_n, DEPOT_MAP[dep_n], prod_n, PRODUCT_MAP[prod_n])
                for bdc_n in fetch_bdcs
                for dep_n in fetch_depots
                for prod_n in fetch_products
            ]
            total_t = len(tasks)
            results = []
            done    = [0]
            lock    = threading.Lock()

            prog_bar  = st.progress(0, text="Starting fetch…")
            log_box   = st.empty()
            log_lines = []

            def _run(args):
                bdc_n, bdc_id, dep_n, dep_id, prod_n, prod_id = args
                params = {
                    "lngProductId": prod_id,
                    "lngBDCId":     bdc_id,
                    "lngDepotId":   dep_id,
                    "dtpStartDate": start_str,
                    "dtpEndDate":   end_str,
                    "lngUserId":    api_user_id,
                }
                pdf_bytes = _fetch_pdf_bytes(STOCK_TXN_URL, params)
                icon = note = ""
                if pdf_bytes:
                    recs = _parse_stock_transaction_pdf(pdf_bytes)
                    if recs:
                        bal = _get_opening_closing(recs)
                        with lock:
                            results.append({
                                "BDC":          bdc_n,
                                "Depot":        dep_n,
                                "Product":      prod_n,
                                "Opening (LT)": bal["opening"],
                                "Closing (LT)": bal["closing"],
                                "Transactions": len(recs),
                            })
                        icon = "✅"
                        note = f"{prod_n} @ {dep_n[:28]} — open: {bal['opening']:,}  close: {bal['closing']:,}"
                    else:
                        icon = "⚠️"
                        note = f"{prod_n} @ {dep_n[:28]} — PDF ok but no records parsed"
                else:
                    icon = "○"
                    note = f"{prod_n} @ {dep_n[:28]} — no data"

                with lock:
                    done[0] += 1
                    pct = done[0] / total_t
                    prog_bar.progress(pct, text=f"Fetching… {done[0]:,} / {total_t:,} ({pct*100:.0f}%)")
                    log_lines.append(f"{icon} [{bdc_n}] {note}")
                    log_box.markdown(
                        "<div class='fetch-log-box'>" + "<br>".join(log_lines[-15:]) + "</div>",
                        unsafe_allow_html=True,
                    )

            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as ex:
                list(ex.map(_run, tasks))

            prog_bar.progress(1.0, text=f"✅ Complete — {len(results):,} records with data")

            if results:
                stock_df = pd.DataFrame(results).sort_values(["BDC", "Depot", "Product"]).reset_index(drop=True)
                st.session_state.stock_balance_df = stock_df
                st.success(f"✅ Retrieved balance data for **{len(stock_df):,}** BDC/Depot/Product combinations.")
            else:
                st.session_state.stock_balance_df = pd.DataFrame()
                st.warning(
                    "⚠️ No data returned. Verify BDC entity IDs in `BDC_ENTITY_MAP` match the NPA system, "
                    "and that there were transactions in the selected month. "
                    "The Excel will still export with empty Opening/Closing sheets."
                )

        st.markdown("---")

        # ── Generate Excel ─────────────────────────────────────────────────────
        st.markdown("#### Step 2 — Generate & Download Excel Report")
        st.markdown("Click below to generate and download the Excel report with **P1**, **SUMMARY**, **OPENING STOCK BALANCE** and **CLOSING STOCK BALANCE** sheets.")

        stock_df = st.session_state.get("stock_balance_df", None)

        if stock_df is not None and not stock_df.empty:
            st.caption(
                f"Stock data loaded: **{len(stock_df):,}** records · "
                f"Opening total: **{stock_df['Opening (LT)'].sum():,.0f} LT** · "
                f"Closing total: **{stock_df['Closing (LT)'].sum():,.0f} LT**"
            )
        else:
            st.caption("ℹ️ Stock balance not yet fetched — Opening & Closing sheets in the Excel will be empty.")

        if st.button("📥 Generate Excel Report", type="primary"):
            with st.spinner("Building Excel file…"):
                if stock_df is not None and not stock_df.empty:
                    opening_df = stock_df[["BDC", "Depot", "Product", "Opening (LT)"]].copy()
                    closing_df = stock_df[["BDC", "Depot", "Product", "Closing (LT)"]].copy()
                else:
                    opening_df = pd.DataFrame(columns=["BDC", "Depot", "Product", "Opening (LT)"])
                    closing_df = pd.DataFrame(columns=["BDC", "Depot", "Product", "Closing (LT)"])

                excel_buf = generate_excel(tables, summary_df, month_label, opening_df, closing_df)

            fname = f"Report_{months[sel_month]}_{sel_year}.xlsx"
            st.download_button(
                label=f"⬇️ Download {fname}",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )