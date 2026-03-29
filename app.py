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
    background: linear-gradient(135deg, #e0f7fa, #e6e6fa, #f3e5f5);
}
.title {
    text-align:center;
    font-size:42px;
    font-weight:bold;
    color:#4A148C;
}
.card {
    background:white;
    padding:20px;
    border-radius:15px;
    box-shadow:0px 5px 15px rgba(0,0,0,0.1);
    margin-bottom:20px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>🦄 Smart Dataset Cleaner</div>", unsafe_allow_html=True)
st.markdown("<div class='card'>Upload → Clean → Analyze → Download 🚀</div>", unsafe_allow_html=True)

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

# ---------------- MAIN ----------------
if file:
    try:
        df_original = pd.read_excel(file)
        df = df_original.copy()

        # CLEAN COLUMN NAMES
        df.columns = df.columns.str.strip().str.upper().str.replace("\n","")

        st.write("📌 Columns detected:", list(df.columns))

        # AUTO DETECT COLUMNS
        def find_column(names):
            for col in df.columns:
                for n in names:
                    if n in col:
                        return col
            return None

        desc = find_column(["DESCRIPTION"])
        qty = find_column(["QUANTITY","QTY"])
        price = find_column(["PRICE","VALUE","INV"])
        curr = find_column(["CURRENCY","CURR"])
        unit = find_column(["UNIT"])
        exporter = find_column(["EXPORTER"])
        consignee = find_column(["CONSIGNEE"])
        country = find_column(["COUNTRY"])

        st.write("🧠 Column Mapping:", {
            "Description": desc,
            "Quantity": qty,
            "Price": price,
            "Currency": curr,
            "Unit": unit
        })

        if not desc or not qty or not price:
            st.error("❌ Required columns not found")
            st.stop()

        # CLEAN DATA
        df[qty] = pd.to_numeric(df[qty], errors="coerce").fillna(0)
        df[price] = pd.to_numeric(df[price], errors="coerce").fillna(0)

        # STANDARD UNIT
        def standardize(row):
            u = str(row.get(unit, "")).lower()
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

        # CLASSIFICATION
        def classify(d):
            d = str(d).lower()
            if dataset_type == "vaccine":
                return "Pediatric Dose" if "pediatric" in d else "Adult Dose"
            return "General"

        df["Classification"] = df[desc].apply(classify)

        # USD PRICE
        def convert(row):
            c = str(row.get(curr, "USD")).strip()
            rate = exchange_rates.get(c, 1)
            return row[price] * rate

        df["USD Price"] = df.apply(convert, axis=1)

        # UNIT PRICE
        df["Unit Price"] = df.apply(
            lambda x: x["USD Price"]/x["Std Qty"] if x["Std Qty"] > 0 else 0,
            axis=1
        )

        # CREATE EXCEL
        wb = Workbook()

        # ORIGINAL
        ws1 = wb.active
        ws1.title = "Original"
        for r in dataframe_to_rows(df_original, index=False, header=True):
            ws1.append(r)

        # CLEANED
        ws2 = wb.create_sheet("Cleaned")
        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # HEADER STYLE
        for i, col in enumerate(df.columns, 1):
            cell = ws2.cell(1, i)
            if col in ["Classification","USD Price","Unit Price","Std Qty","Std Unit"]:
                cell.fill = orange
            cell.font = Font(bold=True)

        desc_i = list(df.columns).index(desc)+1
        qty_i = list(df.columns).index(qty)+1

        # COLOR RULES
        for i, row in df.iterrows():
            r = i+2
            d = str(row[desc]).lower()

            desc_cell = ws2.cell(r, desc_i)
            qty_cell = ws2.cell(r, qty_i)

            if dataset_type == "vaccine" and ("free of cost" in d or "foc" in d):
                desc_cell.fill = dark_red
                qty_cell.fill = dark_red

        # ANALYSIS
        ws3 = wb.create_sheet("Analysis")

        def top10_avg(col):
            if col:
                return df.groupby(col)["Unit Price"].mean().nlargest(10).reset_index()
            return pd.DataFrame()

        def top10_qty(col):
            if col:
                return df.groupby(col)["Std Qty"].sum().nlargest(10).reset_index()
            return pd.DataFrame()

        p1 = top10_avg(country)
        p2 = top10_qty(country)
        p3 = top10_avg(consignee)
        p4 = top10_avg(exporter)
        p5 = top10_qty(exporter)
        p6 = df.groupby("Classification")["Std Qty"].sum().reset_index()

        def write(ws, title, data):
            ws.append([title])
            start = ws.max_row
            for r in dataframe_to_rows(data, index=False, header=True):
                ws.append(r)
            ws.append([])
            return start, len(data)

        r1,n1 = write(ws3,"Country vs Unit Price",p1)
        r2,n2 = write(ws3,"Country vs Quantity",p2)
        r3,n3 = write(ws3,"Consignee vs Unit Price",p3)
        r4,n4 = write(ws3,"Exporter vs Unit Price",p4)
        r5,n5 = write(ws3,"Exporter vs Quantity",p5)
        r6,n6 = write(ws3,"Classification vs Quantity",p6)

        def chart(ws, row, n, title):
            if n == 0:
                return
            c = BarChart()
            c.title = title
            data = Reference(ws, min_col=2, min_row=row+1, max_row=row+n)
            cats = Reference(ws, min_col=1, min_row=row+2, max_row=row+n)
            c.add_data(data)
            c.set_categories(cats)
            ws.add_chart(c, f"E{row}")

        chart(ws3,r1,n1,"Country Price")
        chart(ws3,r2,n2,"Country Qty")
        chart(ws3,r3,n3,"Consignee Price")
        chart(ws3,r4,n4,"Exporter Price")
        chart(ws3,r5,n5,"Exporter Qty")
        chart(ws3,r6,n6,"Classification Qty")

        # DASHBOARD
        ws4 = wb.create_sheet("Dashboard")
        ws4["A1"] = "📊 DASHBOARD"
        ws4["A3"] = f"Total Records: {len(df)}"
        ws4["A4"] = f"Total Quantity: {df['Std Qty'].sum()}"
        ws4["A5"] = f"Total Value (USD): {df['USD Price'].sum()}"

        # SAVE
        output = "final_output.xlsx"
        wb.save(output)

        st.success("✅ File Processed Successfully")

        with open(output, "rb") as f:
            st.download_button("⬇ Download Excel", f, file_name="final_output.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")