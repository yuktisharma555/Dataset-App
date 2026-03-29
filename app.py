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

# ---------------- COLORS ----------------
dark_red = PatternFill("solid", fgColor="8B0000")
light_red = PatternFill("solid", fgColor="FF9999")
light_orange = PatternFill("solid", fgColor="FFD580")

# ---------------- MAIN ----------------
if file:

    try:
        df = pd.read_excel(file)

        # CLEAN COLUMN NAMES
        df.columns = df.columns.str.upper().str.strip()

        st.write("Detected Columns:", df.columns.tolist())

        # -------- EXACT COLUMN MAPPING --------
        desc = "GOODS DESCRIPTION"
        qty = "QUANTITY"
        unit = "UNIT"
        price = "ITEM PRICE_INV"
        total_price = "TOTAL PRICE_INV_FC"
        country = "COUNTRY"
        exporter = "EXPORTER"
        consignee = "CONSIGNEE NAME"
        std_qty_col = "STD QUANTITY"
        std_unit_col = "STD UNIT"

        # -------- SAFETY CHECK --------
        required = [desc, qty, price, country]
        missing = [c for c in required if c not in df.columns]

        if missing:
            st.error(f"❌ Missing columns: {missing}")
            st.stop()

        # -------- CLEAN TYPES --------
        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")
        df[total_price] = pd.to_numeric(df[total_price], errors="coerce")

        # -------- STANDARD QUANTITY --------
        if std_qty_col in df.columns:
            df["Std Quantity"] = pd.to_numeric(df[std_qty_col], errors="coerce")
        else:
            def convert_std(q, u):
                u = str(u).lower()
                if "nos" in u or "pcs" in u or "vial" in u:
                    return q
                elif "ml" in u:
                    return q / 10
                elif "gm" in u:
                    return q / 100
                return q

            df["Std Quantity"] = df.apply(lambda x: convert_std(x[qty], x[unit]), axis=1)

        # -------- UNIT PRICE --------
        df["Unit Price"] = df[total_price] / df["Std Quantity"]

        # -------- CLASSIFICATION --------
        def classify(d):
            d = str(d).lower()
            if dataset_type == "vaccine":
                return "Pediatric" if "pediatric" in d else "Adult"
            return "Other"

        df["Classification"] = df[desc].apply(classify)

        # -------- CREATE EXCEL --------
        wb = Workbook()

        # ORIGINAL SHEET
        ws1 = wb.active
        ws1.title = "Original"

        for r in dataframe_to_rows(df, index=False, header=True):
            ws1.append(r)

        # CLEANED SHEET
        ws2 = wb.create_sheet("Cleaned")

        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # -------- COLOR NEW HEADERS --------
        for i, col in enumerate(df.columns, 1):
            if col in ["Std Quantity", "Unit Price", "Classification"]:
                ws2.cell(1, i).fill = light_orange
                ws2.cell(1, i).font = Font(bold=True)

        desc_idx = list(df.columns).index(desc) + 1

        # -------- ROW COLOR RULES --------
        for i, row in df.iterrows():
            r = i + 2
            d = str(row[desc]).lower()

            if "free of cost" in d or "foc" in d:
                ws2.cell(r, desc_idx).fill = dark_red

            if pd.notna(row[qty]) and row[qty] < 1:
                ws2.cell(r, desc_idx).fill = dark_red

            if pd.notna(row[price]) and row[price] == 0:
                ws2.cell(r, desc_idx).fill = dark_red

            if any(x in d for x in ["sample","standard","r&d","impurity"]):
                ws2.cell(r, desc_idx).fill = dark_red

            if dataset_type == "medicine" and "api" in str(row["Classification"]).lower():
                ws2.cell(r, desc_idx).fill = light_red

        # -------- PIVOT TABLE DATA --------
        def pivot(col, val):
            return df.groupby(col)[val].sum().nlargest(10).reset_index()

        p1 = pivot(country, "Unit Price")
        p2 = pivot(country, "Std Quantity")
        p3 = pivot(consignee, "Unit Price")
        p4 = pivot(exporter, "Unit Price")
        p5 = pivot(exporter, "Std Quantity")
        p6 = pivot("Classification", "Std Quantity")

        # -------- ANALYSIS SHEET --------
        ws3 = wb.create_sheet("Analysis")

        def add_table(title, data, start_row):
            ws3.cell(start_row, 1, title)

            for i, r in enumerate(dataframe_to_rows(data, index=False, header=True)):
                for j, val in enumerate(r):
                    ws3.cell(start_row + i + 1, j + 1, val)

            chart = BarChart()
            chart.title = title

            data_ref = Reference(ws3, min_col=2, min_row=start_row+2, max_row=start_row+11)
            cat_ref = Reference(ws3, min_col=1, min_row=start_row+3, max_row=start_row+11)

            chart.add_data(data_ref)
            chart.set_categories(cat_ref)

            ws3.add_chart(chart, f"E{start_row}")

            return start_row + 15

        row = 1
        row = add_table("Country vs Unit Price", p1, row)
        row = add_table("Country vs Quantity", p2, row)
        row = add_table("Consignee vs Unit Price", p3, row)
        row = add_table("Exporter vs Unit Price", p4, row)
        row = add_table("Exporter vs Quantity", p5, row)
        row = add_table("Classification vs Quantity", p6, row)

        # -------- DASHBOARD --------
        ws4 = wb.create_sheet("Dashboard")

        ws4["A1"] = "DATASET DASHBOARD"
        ws4["A3"] = "Total Value"
        ws4["B3"] = df[total_price].sum()

        ws4["A4"] = "Total Quantity"
        ws4["B4"] = df["Std Quantity"].sum()

        ws4["A5"] = "Top Country"
        ws4["B5"] = p1.iloc[0,0] if not p1.empty else "N/A"

        # SAVE FILE
        wb.save("final.xlsx")

        st.success("✅ Done Successfully")

        with open("final.xlsx", "rb") as f:
            st.download_button("⬇ Download Excel", f, "final.xlsx")

    except Exception as e:
        st.error(str(e))