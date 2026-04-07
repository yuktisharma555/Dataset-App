import streamlit as st
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# ---------------- UI ----------------
st.set_page_config(layout="wide")

st.markdown("<h1 style='text-align:center;'>📊 Smart Dataset Cleaner Pro</h1>", unsafe_allow_html=True)

file = st.file_uploader("Upload Excel", type=["xlsx"])
dataset_type = st.selectbox("Dataset Type", ["medicine","vaccine","testkit"])

# COLORS
dark_red = PatternFill("solid", fgColor="8B0000")
light_orange = PatternFill("solid", fgColor="FFD580")

if file:
    try:
        df = pd.read_excel(file, header=5)
        df.columns = df.columns.astype(str).str.upper().str.strip()

        desc = "GOODS DESCRIPTION"
        qty = "QUANTITY"
        price = "ITEM PRICE_INV"
        currency_col = "CURRENCY"

        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df.get(price, 0), errors="coerce")

        # ---------------- USD (UNIT PRICE) ----------------
        exchange_rates = {
            "ZAR": 0.0628, "THB": 0.0322, "CAD": 0.7334, "INR": 0.0109,
            "MXN": 0.058, "RUB": 0.0129, "GBP": 1.3455, "EUR": 1.1821,
            "AED": 0.2722, "USD": 1
        }

        def convert_unit_price(row):
            curr = str(row.get(currency_col, "USD")).upper()
            rate = exchange_rates.get(curr, 1)
            return row[price] * rate if pd.notna(row[price]) else 0

        df["Item Price (USD)"] = df.apply(convert_unit_price, axis=1)

        # POSITION AFTER CURRENCY
        if currency_col in df.columns:
            cols = list(df.columns)
            cols.remove("Item Price (USD)")
            idx = cols.index(currency_col)
            cols.insert(idx + 1, "Item Price (USD)")
            df = df[cols]

        # ---------------- TEST KIT LOGIC ----------------
        if dataset_type == "testkit":

            def extract_number(text):
                text = str(text).lower()
                match = re.search(r'(\d+)\s*(t|test|tests|rapid)', text)
                return int(match.group(1)) if match else 1

            df["Standard Quantity"] = df[desc].apply(extract_number)
            df["Adjusted Quantity"] = df["Standard Quantity"] * df[qty]

            # POSITIONING
            cols = list(df.columns)
            cols.remove("Standard Quantity")
            cols.remove("Adjusted Quantity")

            q_idx = cols.index(qty)

            cols.insert(q_idx, "Standard Quantity")
            cols.insert(q_idx + 2, "Adjusted Quantity")

            df = df[cols]

        # ---------------- EXCEL ----------------
        wb = Workbook()
        ws = wb.active
        ws.title = "Cleaned Data"

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # HEADER STYLING
        for i, col in enumerate(df.columns, 1):
            if col in ["Standard Quantity", "Adjusted Quantity"]:
                ws.cell(1, i).fill = light_orange
                ws.cell(1, i).font = Font(bold=True)

        # ---------------- ROW LEVEL FORMATTING ----------------
        for row in range(2, ws.max_row + 1):
            text = str(ws.cell(row, df.columns.get_loc(desc)+1).value).lower()

            # Vaccine FREE SAMPLE
            if dataset_type == "vaccine":
                if "free sample" in text or "free quantity" in text:
                    ws.cell(row, df.columns.get_loc(desc)+1).fill = dark_red

        wb.save("final.xlsx")

        st.success("✅ Done")

        with open("final.xlsx", "rb") as f:
            st.download_button("Download", f, "final.xlsx")

    except Exception as e:
        st.error(e)