import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

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
        # ORIGINAL DATA
        original_df = pd.read_excel(file, header=5)

        # CLEAN COPY
        df = original_df.copy()
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
        date_col = "DATE"

        # -------- YEAR EXTRACTION --------
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
            df["YEAR"] = df[date_col].dt.year

            cols = list(df.columns)
            cols.remove("YEAR")
            date_index = cols.index(date_col)
            cols.insert(date_index, "YEAR")
            df = df[cols]

        # -------- CLEAN TYPES --------
        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[total_price] = pd.to_numeric(df[total_price], errors="coerce")

        # STANDARD QUANTITY
        if std_qty_col in df.columns:
            df["Std Quantity"] = pd.to_numeric(df[std_qty_col], errors="coerce")
        else:
            df["Std Quantity"] = df[qty]

        # UNIT PRICE
        df["Unit Price"] = df.apply(
            lambda x: x[total_price] / x["Std Quantity"] if pd.notna(x["Std Quantity"]) and x["Std Quantity"] != 0 else 0,
            axis=1
        )

        # -------- CLASSIFICATION --------
        def classify(d):
            d = str(d).lower()

            if dataset_type == "vaccine":
                return "Pediatric" if "pediatric" in d else "Adult"

            elif dataset_type == "medicine":
                return "API" if "api" in d else "Formulation"

            elif dataset_type == "testkit":
                if any(x in d for x in ["research", "r&d"]):
                    return "Research"
                elif any(x in d for x in ["control", "qc", "quality"]):
                    return "Quality Control"
                elif any(x in d for x in ["screening", "rapid"]):
                    return "Screening"
                else:
                    return "Diagnostic"

            return "Diagnostic"

        df["Classification"] = df[desc].apply(classify)

        # MOVE CLASSIFICATION
        cols = list(df.columns)
        cols.remove("Classification")
        desc_index = cols.index(desc)
        cols.insert(desc_index + 1, "Classification")
        df = df[cols]

        # ---------------- EXCEL ----------------
        wb = Workbook()

        # ORIGINAL SHEET
        ws0 = wb.active
        ws0.title = "Original"
        for r in dataframe_to_rows(original_df, index=False, header=True):
            ws0.append(r)

        # CLEANED SHEET
        ws1 = wb.create_sheet("Cleaned")
        for r in dataframe_to_rows(df, index=False, header=True):
            ws1.append(r)

        # HEADER COLOR
        for i, col in enumerate(df.columns, 1):
            if col in ["Std Quantity", "Unit Price", "Classification", "YEAR"]:
                ws1.cell(1, i).fill = light_orange
                ws1.cell(1, i).font = Font(bold=True)

        # ---------------- ANALYSIS ----------------
        ws2 = wb.create_sheet("Analysis")

        def pivot(col, val):
            data = df.groupby(col)[val].sum().reset_index()
            total = data[val].sum()
            data[val] = (data[val] / total) * 100  # % conversion
            return data.sort_values(by=val, ascending=False).head(10)

        tables = [
            ("Country vs Unit Price %", pivot(country, "Unit Price")),
            ("Country vs Quantity %", pivot(country, "Std Quantity")),
            ("Consignee vs Unit Price %", pivot(consignee, "Unit Price")),
            ("Exporter vs Unit Price %", pivot(exporter, "Unit Price")),
            ("Exporter vs Quantity %", pivot(exporter, "Std Quantity")),
            ("Classification vs Quantity %", pivot("Classification", "Std Quantity"))
        ]

        row = 1
        for title, data in tables:
            ws2.cell(row, 1, title)

            start = row + 1

            for i, rdata in enumerate(dataframe_to_rows(data, index=False, header=True)):
                for j, val in enumerate(rdata):
                    ws2.cell(start + i, j + 1, val)

            max_row = start + len(data)

            chart = BarChart()
            chart.title = title
            chart.y_axis.title = "Percentage (%)"
            chart.x_axis.title = "Category"

            data_ref = Reference(ws2, min_col=2, min_row=start, max_row=max_row)
            cat_ref = Reference(ws2, min_col=1, min_row=start+1, max_row=max_row)

            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cat_ref)

            # DATA LABELS
            chart.dLbls = DataLabelList()
            chart.dLbls.showVal = True

            ws2.add_chart(chart, f"E{row}")

            row += 15

        # ---------------- DASHBOARD ----------------
        ws3 = wb.create_sheet("Dashboard")

        ws3["A1"] = "DATASET DASHBOARD"
        ws3["A3"] = "Total Value"
        ws3["B3"] = df[total_price].sum()

        ws3["A4"] = "Total Quantity"
        ws3["B4"] = df["Std Quantity"].sum()

        ws3["A5"] = "Top Country"
        ws3["B5"] = tables[0][1].iloc[0,0] if not tables[0][1].empty else "N/A"

        wb.save("final.xlsx")

        st.success("✅ Final File Ready with Professional Charts")

        with open("final.xlsx", "rb") as f:
            st.download_button("⬇ Download Excel", f, "final.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {e}")