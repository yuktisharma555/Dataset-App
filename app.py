import streamlit as st
import pandas as pd
import xlwings as xw
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(layout="wide")
st.title("📊 Dataset Cleaner + Real Pivot Generator")

file = st.file_uploader("Upload Excel", type=["xlsx"])

# ---------------- CLEAN + SAVE ----------------
def create_clean_excel(df, file_path):

    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned"

    # leave first 5 rows blank → header at row 6
    for _ in range(5):
        ws.append([])

    # write dataframe (header will go to row 6)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    wb.save(file_path)


# ---------------- XLWINGS PIVOT ----------------
def create_pivots(file_path):

    app = xw.App(visible=False)
    wb = app.books.open(file_path)

    ws = wb.sheets['Cleaned']

    # 🔥 YOUR FIXED RANGE
    last_row = ws.range("A6").expand("down").last_cell.row
    source_range = f"A6:AJ{last_row}"

    try:
        pivot_sheet = wb.sheets['Pivots']
        pivot_sheet.clear()
    except:
        pivot_sheet = wb.sheets.add("Pivots")

    pivot_cache = wb.api.PivotCaches().Create(
        SourceType=1,
        SourceData=ws.range(source_range).api
    )

    # -------- PIVOT 1 --------
    pt1 = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range("A3").api,
        TableName="PivotCountryPrice"
    )
    pt1.PivotFields("COUNTRY").Orientation = 1
    pt1.AddDataField(pt1.PivotFields("UNIT PRICE_INR"), "Avg Unit Price", -4106)

    # -------- PIVOT 2 --------
    pt2 = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range("H3").api,
        TableName="PivotCountryQty"
    )
    pt2.PivotFields("COUNTRY").Orientation = 1
    pt2.AddDataField(pt2.PivotFields("STD QUANTITY"), "Total Qty", -4157)

    # -------- PIVOT 3 --------
    pt3 = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range("A20").api,
        TableName="PivotExporterPrice"
    )
    pt3.PivotFields("EXPORTER").Orientation = 1
    pt3.AddDataField(pt3.PivotFields("UNIT PRICE_INR"), "Avg Price", -4106)

    # -------- PIVOT 4 --------
    pt4 = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range("H20").api,
        TableName="PivotExporterQty"
    )
    pt4.PivotFields("EXPORTER").Orientation = 1
    pt4.AddDataField(pt4.PivotFields("STD QUANTITY"), "Total Qty", -4157)

    # -------- PIVOT 5 --------
    pt5 = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range("A37").api,
        TableName="PivotConsignee"
    )
    pt5.PivotFields("CONSIGNEE NAME").Orientation = 1
    pt5.AddDataField(pt5.PivotFields("UNIT PRICE_INR"), "Avg Price", -4106)

    # -------- PIVOT 6 --------
    pt6 = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range("H37").api,
        TableName="PivotClassification"
    )
    pt6.PivotFields("CLASSIFICATION").Orientation = 1
    pt6.AddDataField(pt6.PivotFields("STD QUANTITY"), "Total Qty", -4157)

    wb.save()
    wb.close()
    app.quit()


# ---------------- MAIN ----------------
if file:
    try:
        # read correctly (row 6 header)
        df = pd.read_excel(file, header=5)
        df.columns = df.columns.str.upper().str.strip()

        st.success("✅ File Loaded")

        # -------- ADD CLASSIFICATION --------
        def classify(d):
            d = str(d).lower()
            if "pediatric" in d:
                return "Pediatric"
            return "Adult"

        df["CLASSIFICATION"] = df["GOODS DESCRIPTION"].apply(classify)

        # move classification after description
        cols = list(df.columns)
        cols.remove("CLASSIFICATION")
        desc_index = cols.index("GOODS DESCRIPTION")
        cols.insert(desc_index + 1, "CLASSIFICATION")
        df = df[cols]

        file_path = "final.xlsx"

        # save cleaned
        create_clean_excel(df, file_path)

        # create pivots
        create_pivots(file_path)

        st.success("✅ Pivot Tables Created Successfully")

        with open(file_path, "rb") as f:
            st.download_button("⬇ Download Final Excel", f, "final.xlsx")

    except Exception as e:
        st.error(str(e))