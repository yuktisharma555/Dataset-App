import streamlit as st
import pandas as pd
import re
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
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align:center;'>📊 Smart Dataset Cleaner Pro</h1>", unsafe_allow_html=True)

file = st.file_uploader("Upload Excel", type=["xlsx"])
dataset_type = st.selectbox("Dataset Type", ["medicine","vaccine","testkit"])

# ---------------- COLORS ----------------
dark_red = PatternFill("solid", fgColor="8B0000")
light_red = PatternFill("solid", fgColor="FF9999")
light_orange = PatternFill("solid", fgColor="FFD580")

# ---------------- MAIN ----------------
if file:
    try:
        df = pd.read_excel(file, header=5)
        df.columns = df.columns.astype(str).str.upper().str.strip()

        st.success("✅ Header loaded correctly")

        # COLUMN MAPPING
        desc = "GOODS DESCRIPTION"
        qty = "QUANTITY"
        price = "ITEM PRICE_INV"
        total_price = "TOTAL PRICE_INV_FC"
        country = "COUNTRY"
        exporter = "EXPORTER"
        consignee = "CONSIGNEE NAME"
        std_qty_col = "STD QUANTITY"
        currency_col = "CURRENCY"

        # SAFETY CHECK
        required = [desc, qty, total_price, country]
        missing = [c for c in required if c not in df.columns]

        if missing:
            st.error(f"❌ Missing columns: {missing}")
            st.stop()

        # CLEAN TYPES
        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df.get(price, 0), errors="coerce")
        df[total_price] = pd.to_numeric(df[total_price], errors="coerce")

        # ---------------- STD QUANTITY LOGIC ----------------
        def extract_tablets(desc_text):
            text = str(desc_text).lower()
            match = re.search(r'(\d+)\s*[x\*]\s*(\d+)\s*(tab|tabs|tablet|tablets)', text)
            if match:
                return int(match.group(1)) * int(match.group(2))
            return None

        if dataset_type == "medicine":
            df["Std Quantity"] = df.apply(
                lambda x: extract_tablets(x[desc]) if extract_tablets(x[desc]) else x[qty],
                axis=1
            )
        else:
            if std_qty_col in df.columns:
                df["Std Quantity"] = pd.to_numeric(df[std_qty_col], errors="coerce")
            else:
                df["Std Quantity"] = df[qty]

        # ---------------- UNIT PRICE ----------------
        df["Unit Price"] = df.apply(
            lambda x: x[total_price] / x["Std Quantity"] if pd.notna(x["Std Quantity"]) and x["Std Quantity"] != 0 else 0,
            axis=1
        )

        # ---------------- USD CONVERSION ----------------
        exchange_rates = {
            "ZAR": 0.0628, "THB": 0.0322, "CAD": 0.7334, "INR": 0.0109,
            "MXN": 0.058, "RUB": 0.0129, "GBP": 1.3455, "EUR": 1.1821,
            "AED": 0.2722, "USD": 1
        }

        def convert_to_usd(row):
            curr = str(row.get(currency_col, "USD")).upper()
            rate = exchange_rates.get(curr, 1)
            return row[total_price] * rate if pd.notna(row[total_price]) else 0

        df["Item Price (USD)"] = df.apply(convert_to_usd, axis=1)

        # ---------------- CLASSIFICATION ----------------
        def classify(d):
            d = str(d).lower()

            if dataset_type == "vaccine":
                return "Pediatric" if "pediatric" in d else "Adult"

            elif dataset_type == "medicine":
                return "API" if "api" in d else "Formulation"

            elif dataset_type == "testkit":
                if "card" in d:
                    return "Card"
                elif "strip" in d or "dipstick" in d:
                    return "Strip"
                elif "test kit" in d or "test" in d:
                    return "Full Testing Kit"
                else:
                    return "Other"

            return "Other"

        df["Classification"] = df[desc].apply(classify)

        # MOVE CLASSIFICATION
        cols = list(df.columns)
        cols.remove("Classification")
        desc_index = cols.index(desc)
        cols.insert(desc_index + 1, "Classification")
        df = df[cols]

        # ---------------- EXCEL ----------------
        wb = Workbook()

        ws1 = wb.active
        ws1.title = "Cleaned"

        for r in dataframe_to_rows(df, index=False, header=True):
            ws1.append(r)

        # HEADER COLOR
        for i, col in enumerate(df.columns, 1):
            if col in ["Std Quantity", "Unit Price", "Classification", "Item Price (USD)"]:
                ws1.cell(1, i).fill = light_orange
                ws1.cell(1, i).font = Font(bold=True)

        desc_idx = list(df.columns).index(desc) + 1

        # ---------------- ROW COLOR RULES ----------------
        for i, row in df.iterrows():
            r = i + 2
            d = str(row[desc]).lower()

            if dataset_type == "testkit":

                if ("kit" in d or "testing kit" in d) and any(x in d for x in ["blood", "plasma"]):
                    pass

                elif any(x in d for x in [
                    "raw material", "row material",
                    "biological substance",
                    "plasma", "blood", "serum", "sample", "samples",
                    "free supply",
                    "for testing purpose",
                    "r&d purpose"
                ]):
                    ws1.cell(r, desc_idx).fill = dark_red

            else:
                if "free of cost" in d or "foc" in d:
                    ws1.cell(r, desc_idx).fill = dark_red

                if pd.notna(row[qty]) and row[qty] < 1:
                    ws1.cell(r, desc_idx).fill = dark_red

                if pd.notna(row[price]) and row[price] == 0:
                    ws1.cell(r, desc_idx).fill = dark_red

                if any(x in d for x in ["sample","standard","r&d","impurity"]):
                    ws1.cell(r, desc_idx).fill = dark_red

                if dataset_type == "medicine" and "api" in str(row["Classification"]).lower():
                    ws1.cell(r, desc_idx).fill = light_red

        # ---------------- DASHBOARD ----------------
        ws3 = wb.create_sheet("Dashboard")

        ws3["A1"] = "DATASET DASHBOARD"
        ws3["A3"] = "Total Value"
        ws3["B3"] = df[total_price].sum()

        ws3["A4"] = "Total Quantity"
        ws3["B4"] = df["Std Quantity"].sum()

        ws3["A5"] = "Top Country"
        ws3["B5"] = df.groupby(country)["Unit Price"].sum().idxmax()

        wb.save("final.xlsx")

        st.success("✅ File Processed Successfully")

        with open("final.xlsx", "rb") as f:
            st.download_button("⬇ Download Excel", f, "final.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {e}")