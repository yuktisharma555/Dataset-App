import streamlit as st
import pandas as pd
import tempfile
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# ---------------- UI ----------------
st.set_page_config(page_title="Dataset Analyzer", layout="wide")

st.markdown("""
<style>
.stApp {
    background-color: #E3F2FD;
}
h1 {
    text-align:center;
    color:#1565C0;
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>📊 Smart Dataset Cleaner</h1>", unsafe_allow_html=True)

file = st.file_uploader("📂 Upload Dataset", type=["xlsx"])
dataset_type = st.selectbox("📊 Dataset Type", ["medicine","vaccine","testkit"])

# ---------------- COLORS ----------------
dark_red = PatternFill("solid", fgColor="8B0000")
light_red = PatternFill("solid", fgColor="FF9999")
orange = PatternFill("solid", fgColor="FFD580")

# ---------------- EXCHANGE ----------------
exchange_rates = {
    "ZAR":0.0628,"THB":0.0322,"CAD":0.7334,"INR":0.0109,
    "MXN":0.058,"RUB":0.0129,"GBP":1.3455,"EUR":1.1821,
    "AED":0.2722,"USD":1
}

# ---------------- CLASSIFICATION ----------------
def classify(desc):
    d = str(desc).lower()

    if dataset_type == "medicine":
        if any(x in d for x in ["lamivudine","efavirenz","emtricitabine","dolutegravir"]):
            return "FFP Combination"
        elif any(x in d for x in ["tablet","capsule","vial","bottle"]):
            return "FFP Plain"
        else:
            return "API"

    elif dataset_type == "vaccine":
        return "Pediatric Dose" if "pediatric" in d else "Adult Dose"

    else:
        if "strip" in d: return "Strip"
        if "card" in d: return "Card"
        return "Kit"

# ---------------- MAIN ----------------
if file:
    try:
        # SAVE TEMP FILE
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(file.read())
            temp_path = tmp.name

        # LOAD ORIGINAL (PRESERVE FORMAT)
        wb = load_workbook(temp_path)
        ws_original = wb.active
        ws_original.title = "Original"

        # READ DATA
        df = pd.read_excel(temp_path)

        df.columns = df.columns.str.upper().str.strip()

        # FIND COLUMNS
        def find(keys):
            for c in df.columns:
                for k in keys:
                    if k in c:
                        return c
            return None

        desc = find(["DESCRIPTION"])
        qty = find(["QTY","QUANTITY"])
        price = find(["PRICE"])
        curr = find(["CURR"])
        unit = find(["UNIT"])
        exporter = find(["EXPORTER"])
        consignee = find(["CONSIGNEE"])
        country = find(["COUNTRY"])

        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        # ---------------- STANDARD UNIT ----------------
        def standardize(row):
            u = str(row[unit]).lower()
            q = row[qty]

            if pd.isna(q):
                return 0, "NOS"

            if "vial" in u:
                return q, "VIAL"
            elif "nos" in u or "pcs" in u:
                return q, "NOS"
            elif "ml" in u:
                return q, "ML"
            else:
                return q, "NOS"

        df[["Std Qty","Std Unit"]] = df.apply(
            lambda x: pd.Series(standardize(x)), axis=1
        )

        # ---------------- NEW COLUMNS ----------------
        df["Classification"] = df[desc].apply(classify)

        df["USD Price"] = df.apply(
            lambda x: x[price]*exchange_rates.get(str(x[curr]).strip(),1),
            axis=1
        )

        df["Unit Price"] = df["USD Price"] / df["Std Qty"]
        df["Unit Price"] = df["Unit Price"].fillna(0)

        # ---------------- ANALYSIS ----------------
        def top10_avg(col):
            if col:
                return df.groupby(col)["Unit Price"].mean().nlargest(10).reset_index()
            return None

        def top10_qty(col):
            if col:
                return df.groupby(col)["Std Qty"].sum().nlargest(10).reset_index()
            return None

        p1 = top10_avg(country)
        p2 = top10_qty(country)
        p3 = top10_avg(consignee)
        p4 = top10_avg(exporter)
        p5 = top10_qty(exporter)
        p6 = df.groupby("Classification")["Std Qty"].sum().reset_index()

        # ---------------- CLEANED SHEET ----------------
        ws_clean = wb.create_sheet("Cleaned")

        for r in dataframe_to_rows(df, index=False, header=True):
            ws_clean.append(r)

        # HEADER COLOR
        for i, col in enumerate(df.columns, 1):
            cell = ws_clean.cell(1, i)
            if col in ["Classification","USD Price","Unit Price","Std Qty","Std Unit"]:
                cell.fill = orange
                cell.font = Font(bold=True)
            else:
                cell.font = Font(bold=True)

        desc_i = list(df.columns).index(desc) + 1
        qty_i = list(df.columns).index(qty) + 1

        # ---------------- ROW COLOR ----------------
        for i, row in df.iterrows():
            r = i + 2
            d = str(row[desc]).lower()
            q = row[qty]
            p = row[price]

            desc_cell = ws_clean.cell(r, desc_i)
            qty_cell = ws_clean.cell(r, qty_i)

            if dataset_type == "vaccine" and ("free of cost" in d or "foc" in d):
                desc_cell.fill = dark_red
                qty_cell.fill = dark_red

            elif pd.notna(q) and q < 1:
                desc_cell.fill = dark_red
                qty_cell.fill = dark_red

            elif pd.notna(p) and p == 0:
                desc_cell.fill = orange

            elif any(x in d for x in ["sample","standard","r&d","impurity"]):
                desc_cell.fill = dark_red

            elif dataset_type == "medicine" and "api" in str(row["Classification"]).lower():
                desc_cell.fill = light_red

        # ---------------- ANALYSIS SHEET ----------------
        ws_analysis = wb.create_sheet("Analysis")

        def write(ws, title, data):
            ws.append([title])
            start = ws.max_row

            if data is not None:
                for r in dataframe_to_rows(data, index=False, header=True):
                    ws.append(r)

            ws.append([])
            return start, len(data) if data is not None else 1

        r1,n1 = write(ws_analysis,"Country vs Unit Price",p1)
        r2,n2 = write(ws_analysis,"Country vs Quantity",p2)
        r3,n3 = write(ws_analysis,"Consignee vs Unit Price",p3)
        r4,n4 = write(ws_analysis,"Exporter vs Unit Price",p4)
        r5,n5 = write(ws_analysis,"Exporter vs Quantity",p5)
        r6,n6 = write(ws_analysis,"Classification vs Quantity",p6)

        def chart(ws, row, n, title):
            c = BarChart()
            c.title = title
            data = Reference(ws, min_col=2, min_row=row+1, max_row=row+n)
            cats = Reference(ws, min_col=1, min_row=row+2, max_row=row+n)
            c.add_data(data)
            c.set_categories(cats)
            ws.add_chart(c, f"E{row}")

        chart(ws_analysis,r1,n1,"Country Price")
        chart(ws_analysis,r2,n2,"Country Qty")
        chart(ws_analysis,r3,n3,"Consignee Price")
        chart(ws_analysis,r4,n4,"Exporter Price")
        chart(ws_analysis,r5,n5,"Exporter Qty")
        chart(ws_analysis,r6,n6,"Classification Qty")

        # ---------------- DASHBOARD ----------------
        ws_dash = wb.create_sheet("Dashboard")

        ws_dash["A1"] = "📊 DATASET DASHBOARD"
        ws_dash["A3"] = f"Total Records: {len(df)}"
        ws_dash["A4"] = f"Total Quantity: {df['Std Qty'].sum()}"
        ws_dash["A5"] = f"Total Value (USD): {df['USD Price'].sum()}"

        # SAVE FILE
        output_path = "final_output.xlsx"
        wb.save(output_path)

        st.success("✅ Processing Complete")

        with open(output_path, "rb") as f:
            st.download_button("⬇ Download Excel", f, file_name="final_output.xlsx")

    except Exception as e:
        st.error(str(e))