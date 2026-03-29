import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# ---------------- UI ----------------
st.set_page_config(layout="wide")

st.markdown("""
<style>
body {
    background: linear-gradient(135deg, #dbeafe, #ffffff, #ede9fe);
}
h1 {text-align:center; color:#6b21a8;}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>🦄 Smart Dataset Cleaner Pro</h1>", unsafe_allow_html=True)

file = st.file_uploader("Upload Excel", type=["xlsx"])
dataset_type = st.selectbox("Dataset Type", ["medicine","vaccine","testkit"])

# COLORS
dark_red = PatternFill("solid", fgColor="8B0000")
light_red = PatternFill("solid", fgColor="FF9999")
light_orange = PatternFill("solid", fgColor="FFD580")

# ---------------- MAIN ----------------
if file:
    df = pd.read_excel(file)
    df.columns = df.columns.str.upper().str.strip()

    def find(col):
        for c in df.columns:
            if col in c:
                return c

    desc = find("DESCRIPTION")
    qty = find("QTY")
    price = find("PRICE")
    unit = find("UNIT")
    country = find("COUNTRY")
    exporter = find("EXPORTER")
    consignee = find("CONSIGNEE")

    df[qty] = pd.to_numeric(df[qty], errors="coerce")
    df[price] = pd.to_numeric(df[price], errors="coerce")

    # STANDARD QTY
    def std_qty(q, u):
        u = str(u).lower()
        if "nos" in u or "pcs" in u or "vial" in u:
            return q
        elif "ml" in u:
            return q / 10
        elif "gm" in u:
            return q / 100
        return q

    df["Std Qty (NOS/PCS/VIAL)"] = df.apply(lambda x: std_qty(x[qty], x[unit]), axis=1)

    # UNIT PRICE
    df["Unit Price"] = df[price] / df[qty]

    # CLASSIFICATION
    def classify(d):
        d = str(d).lower()
        if dataset_type == "vaccine":
            return "Pediatric" if "pediatric" in d else "Adult"
        return "Other"

    df["Classification"] = df[desc].apply(classify)

    # ----------- EXCEL -----------
    wb = Workbook()

    # CLEANED SHEET
    ws = wb.active
    ws.title = "Cleaned"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # COLOR NEW HEADERS
    for i, col in enumerate(df.columns,1):
        if col in ["Std Qty (NOS/PCS/VIAL)", "Unit Price", "Classification"]:
            ws.cell(1,i).fill = light_orange

    desc_idx = list(df.columns).index(desc)+1

    # ROW RULES
    for i,row in df.iterrows():
        r = i+2
        d = str(row[desc]).lower()

        if "free of cost" in d or "foc" in d:
            ws.cell(r,desc_idx).fill = dark_red

        if row[qty] < 1:
            ws.cell(r,desc_idx).fill = dark_red

    # -------- PIVOTS (REAL DATA TABLES) --------
    pivot_sheet = wb.create_sheet("Pivots")

    def pivot(data, idx, val):
        p = data.pivot_table(index=idx, values=val, aggfunc="sum").nlargest(10,val)
        return p.reset_index()

    p1 = pivot(df, country, "Unit Price")
    p2 = pivot(df, country, "Std Qty (NOS/PCS/VIAL)")
    p3 = pivot(df, consignee, "Unit Price")
    p4 = pivot(df, exporter, "Unit Price")
    p5 = pivot(df, exporter, "Std Qty (NOS/PCS/VIAL)")
    p6 = pivot(df, "Classification", "Std Qty (NOS/PCS/VIAL)")

    row = 1
    for title, data in [
        ("Country vs Price",p1),
        ("Country vs Qty",p2),
        ("Consignee vs Price",p3),
        ("Exporter vs Price",p4),
        ("Exporter vs Qty",p5),
        ("Classification vs Qty",p6)
    ]:
        pivot_sheet.append([title])
        start = pivot_sheet.max_row+1

        for r in dataframe_to_rows(data,index=False,header=True):
            pivot_sheet.append(r)

        chart = BarChart()
        data_ref = Reference(pivot_sheet,min_col=2,min_row=start,max_row=start+10)
        cat_ref = Reference(pivot_sheet,min_col=1,min_row=start+1,max_row=start+10)

        chart.add_data(data_ref)
        chart.set_categories(cat_ref)
        chart.title = title

        pivot_sheet.add_chart(chart,f"E{start}")
        row += 15

    # -------- DASHBOARD --------
    dash = wb.create_sheet("Dashboard")
    dash["A1"] = "Dataset Dashboard"
    dash["A3"] = "Total Value"
    dash["B3"] = df[price].sum()

    dash["A4"] = "Total Qty"
    dash["B4"] = df["Std Qty (NOS/PCS/VIAL)"].sum()

    dash["A5"] = "Top Country"
    dash["B5"] = p1.iloc[0,0]

    wb.save("final.xlsx")

    st.success("Done ✅")

    with open("final.xlsx","rb") as f:
        st.download_button("Download",f,"final.xlsx")