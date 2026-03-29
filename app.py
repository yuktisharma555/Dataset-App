import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# ---------------- UI ----------------
st.set_page_config(page_title="Dataset Analyzer", layout="wide")

st.markdown("""
<style>
.stApp {
    background: linear-gradient(to right, #E6E6FA, #D8BFD8);
}
.main-title {
    text-align:center;
    font-size:40px;
    color:#4B0082;
    font-weight:bold;
}
.card {
    background-color:white;
    padding:15px;
    border-radius:12px;
    box-shadow:0px 4px 10px rgba(0,0,0,0.1);
    margin-bottom:15px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>📊 Smart Dataset Cleaner</div>", unsafe_allow_html=True)

st.markdown("<div class='card'>Upload your dataset and get cleaned insights instantly 🚀</div>", unsafe_allow_html=True)

file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])
dataset_type = st.selectbox("📊 Dataset Type", ["medicine","vaccine","testkit"])

# ---------------- COLORS ----------------
dark_red = PatternFill("solid", fgColor="8B0000")
light_red = PatternFill("solid", fgColor="FF9999")
orange = PatternFill("solid", fgColor="FFD580")

exchange_rates = {
    "ZAR":0.0628,"THB":0.0322,"CAD":0.7334,"INR":0.0109,
    "MXN":0.058,"RUB":0.0129,"GBP":1.3455,"EUR":1.1821,
    "AED":0.2722,"USD":1
}

if file:
    try:
        # ---------------- READ FILE ----------------
        df_original = pd.read_excel(file)
        df = df_original.copy()

        df.columns = df.columns.str.upper().str.strip()

        # ---------------- COLUMN MAPPING ----------------
        desc = "GOODS DESCRIPTION"
        qty = "QUANTITY"
        price = "ITEM PRICE_INV"
        curr = "CURRENCY"
        unit = "UNIT"

        # ---------------- CLEAN DATA ----------------
        df[qty] = pd.to_numeric(df[qty], errors="coerce").fillna(0)
        df[price] = pd.to_numeric(df[price], errors="coerce").fillna(0)

        # ---------------- STANDARD UNIT ----------------
        def standardize(row):
            u = str(row[unit]).lower()
            q = row[qty]

            if "vial" in u:
                return q, "VIAL"
            elif "nos" in u or "pcs" in u:
                return q, "NOS"
            else:
                return q, "NOS"

        df[["Std Qty","Std Unit"]] = df.apply(
            lambda x: pd.Series(standardize(x)), axis=1
        )

        # ---------------- CLASSIFICATION ----------------
        def classify(d):
            d = str(d).lower()
            if dataset_type == "vaccine":
                return "Pediatric Dose" if "pediatric" in d else "Adult Dose"
            return "General"

        df["Classification"] = df[desc].apply(classify)

        # ---------------- USD PRICE ----------------
        def convert(row):
            c = str(row[curr]).strip()
            rate = exchange_rates.get(c, 1)
            return row[price] * rate

        df["USD Price"] = df.apply(convert, axis=1)

        # ---------------- UNIT PRICE ----------------
        df["Unit Price"] = df.apply(
            lambda x: x["USD Price"]/x["Std Qty"] if x["Std Qty"] > 0 else 0,
            axis=1
        )

        # ---------------- CREATE EXCEL ----------------
        wb = Workbook()

        # ORIGINAL SHEET
        ws1 = wb.active
        ws1.title = "Original"
        for r in dataframe_to_rows(df_original, index=False, header=True):
            ws1.append(r)

        # CLEANED SHEET
        ws2 = wb.create_sheet("Cleaned")
        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # HEADER COLOR
        for i, col in enumerate(df.columns, 1):
            cell = ws2.cell(1, i)
            if col in ["Classification","USD Price","Unit Price","Std Qty","Std Unit"]:
                cell.fill = orange
            cell.font = Font(bold=True)

        desc_i = list(df.columns).index(desc)+1
        qty_i = list(df.columns).index(qty)+1

        # ---------------- COLOR RULES ----------------
        for i, row in df.iterrows():
            r = i+2
            d = str(row[desc]).lower()

            desc_cell = ws2.cell(r, desc_i)
            qty_cell = ws2.cell(r, qty_i)

            # Vaccine FOC highlight
            if dataset_type == "vaccine" and ("free of cost" in d or "foc" in d):
                desc_cell.fill = dark_red
                qty_cell.fill = dark_red

        # ---------------- ANALYSIS ----------------
        ws3 = wb.create_sheet("Analysis")

        def top10_avg(col):
            return df.groupby(col)["Unit Price"].mean().nlargest(10).reset_index()

        def top10_qty(col):
            return df.groupby(col)["Std Qty"].sum().nlargest(10).reset_index()

        p1 = top10_avg("CURRENCY")
        p2 = top10_qty("CURRENCY")

        def write(ws, title, data):
            ws.append([title])
            start = ws.max_row
            for r in dataframe_to_rows(data, index=False, header=True):
                ws.append(r)
            ws.append([])
            return start, len(data)

        r1,n1 = write(ws3,"Currency vs Unit Price",p1)
        r2,n2 = write(ws3,"Currency vs Quantity",p2)

        def chart(ws, row, n, title):
            c = BarChart()
            c.title = title
            data = Reference(ws, min_col=2, min_row=row+1, max_row=row+n)
            cats = Reference(ws, min_col=1, min_row=row+2, max_row=row+n)
            c.add_data(data)
            c.set_categories(cats)
            ws.add_chart(c, f"E{row}")

        chart(ws3,r1,n1,"Unit Price")
        chart(ws3,r2,n2,"Quantity")

        # ---------------- DASHBOARD ----------------
        ws4 = wb.create_sheet("Dashboard")
        ws4["A1"] = "📊 DASHBOARD"
        ws4["A3"] = f"Total Records: {len(df)}"
        ws4["A4"] = f"Total Quantity: {df['Std Qty'].sum()}"
        ws4["A5"] = f"Total Value: {df['USD Price'].sum()}"

        # SAVE
        output = "final_output.xlsx"
        wb.save(output)

        st.success("✅ File Processed Successfully")

        with open(output, "rb") as f:
            st.download_button("⬇ Download Excel", f, file_name="final_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")