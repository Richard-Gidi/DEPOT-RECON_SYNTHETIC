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


def generate_excel(tables, summary_df, month_label):
    wb = Workbook()
    ws_p1 = wb.active
    ws_p1.title = "P1"
    write_p1_sheet(ws_p1, tables)

    ws_sum = wb.create_sheet("SUMMARY")
    write_summary_sheet(ws_sum, summary_df)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Streamlit UI ───────────────────────────────────────────────────────────────
st.title("📦 Order Request Report Generator")
st.markdown("Upload an Excel file with an **ORDER REQUEST** sheet to generate P1 & Summary reports.")

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
        st.markdown("Click below to generate and download the Excel report with **P1** and **SUMMARY** sheets.")
        if st.button("📥 Generate Excel Report", type="primary"):
            with st.spinner("Building Excel file…"):
                excel_buf = generate_excel(tables, summary_df, month_label)
            fname = f"Report_{months[sel_month]}_{sel_year}.xlsx"
            st.download_button(
                label=f"⬇️ Download {fname}",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )