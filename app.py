import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

st.title("📊 Advanced Dataset Cleaner")

file = st.file_uploader("Upload Excel", type=["xlsx"])
dataset_type = st.selectbox("Dataset Type", ["medicine","vaccine","testkit"])

exchange_rates = {
    "ZAR":0.0628,"THB":0.0322,"CAD":0.7334,"INR":0.0109,
    "MXN":0.058,"RUB":0.0129,"GBP":1.3455,"EUR":1.1821,
    "AED":0.2722,"USD":1
}

# Colors
dark_red = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
light_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

def classify(desc):
    desc = str(desc).lower()

    if dataset_type == "medicine":
        if any(x in desc for x in ["lamivudine","efavirenz","emtricitabine","dolutegravir","rilpivirine"]):
            return "FFP Combination"
        elif any(x in desc for x in ["tablet","capsule","vial","bottle"]):
            return "FFP Plain"
        else:
            return "API"

    elif dataset_type == "vaccine":
        return "Pediatric Dose" if "pediatric" in desc else "Adult Dose"

    else:
        if "strip" in desc:
            return "Strip"
        elif "card" in desc:
            return "Card"
        return "Kit"

if file:
    try:
        original_df = pd.read_excel(file)

        # detect header row
        df = None
        for i in range(6):
            temp = pd.read_excel(file, header=i)
            temp.columns = temp.columns.astype(str).str.upper().str.strip()
            if "GOODS DESCRIPTION" in temp.columns:
                df = temp
                break

        if df is None:
            st.error("❌ Header row not found")
            st.stop()

        df = df.dropna(how="all")

        # dynamic column finder
        def find_col(keys):
            for col in df.columns:
                for k in keys:
                    if k in col:
                        return col
            return None

        desc = find_col(["DESCRIPTION"])
        qty = find_col(["QTY","QUANTITY"])
        price = find_col(["PRICE"])
        curr = find_col(["CURR"])
        unit_col = find_col(["UNIT"])
        exporter_col = find_col(["EXPORTER"])
        consignee_col = find_col(["CONSIGNEE"])
        country_col = find_col(["COUNTRY"])

        if not all([desc, qty, price, curr]):
            st.error("❌ Required columns missing")
            st.stop()

        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        # classification
        df["Classification"] = df[desc].apply(classify)

        # USD conversion
        df["USD Price"] = df.apply(
            lambda x: x[price] * exchange_rates.get(str(x[curr]).strip(),1),
            axis=1
        )

        # standard quantity
        def standardize_qty(q, u):
            u = str(u).lower()
            if "nos" in u or "pcs" in u:
                return q
            elif "vial" in u or "vls" in u:
                return q
            elif "g" in u:
                return q
            return q

        df["Std Qty (NOS/PCS/VIAL)"] = df.apply(
            lambda x: standardize_qty(x[qty], x[unit_col]),
            axis=1
        )

        # -------- ANALYSIS --------
        def top10(group_col, value_col):
            return df.groupby(group_col)[value_col].sum().nlargest(10).reset_index()

        pivot_country_price = top10(country_col, "USD Price") if country_col else None
        pivot_country_qty = top10(country_col, "Std Qty (NOS/PCS/VIAL)") if country_col else None
        pivot_consignee_price = top10(consignee_col, "USD Price") if consignee_col else None
        pivot_exporter_price = top10(exporter_col, "USD Price") if exporter_col else None
        pivot_exporter_qty = top10(exporter_col, "Std Qty (NOS/PCS/VIAL)") if exporter_col else None
        pivot_classification = df.groupby("Classification")["Std Qty (NOS/PCS/VIAL)"].sum().reset_index()

        # -------- CREATE EXCEL --------
        wb = Workbook()

        # ORIGINAL
        ws1 = wb.active
        ws1.title = "Original Data"
        for r in dataframe_to_rows(original_df, index=False, header=True):
            ws1.append(r)

        # CLEANED
        ws2 = wb.create_sheet("Cleaned Data")
        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # header coloring for new columns
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws2.cell(row=1, column=col_idx)
            if col_name in ["Classification","USD Price","Std Qty (NOS/PCS/VIAL)"]:
                cell.fill = orange

        # dynamic description column index
        desc_index = list(df.columns).index(desc) + 1

        # apply row coloring
        for idx, row in df.iterrows():
            desc_val = str(row[desc]).lower()
            qty_val = row[qty]
            price_val = row[price]

            excel_row = idx + 2
            cell = ws2.cell(row=excel_row, column=desc_index)

            if pd.notna(qty_val) and qty_val < 1:
                cell.fill = dark_red

            if pd.notna(price_val) and price_val == 0:
                cell.fill = orange

            if any(x in desc_val for x in ["sample","standard","r&d","impurity"]):
                cell.fill = dark_red

            if dataset_type == "medicine" and "api" in str(row["Classification"]).lower():
                cell.fill = light_red

            if dataset_type == "vaccine" and "free of cost" in desc_val:
                cell.fill = dark_red

        # ANALYSIS SHEET
        ws3 = wb.create_sheet("Analysis")

        def write_table(ws, title, data):
            ws.append([title])
            start_row = ws.max_row

            if data is not None:
                for r in dataframe_to_rows(data, index=False, header=True):
                    ws.append(r)
            else:
                ws.append(["Not Available"])

            ws.append([])
            return start_row

        r1 = write_table(ws3, "Top 10 Countries vs Unit Price", pivot_country_price)
        r2 = write_table(ws3, "Top 10 Countries vs Quantity", pivot_country_qty)
        r3 = write_table(ws3, "Top 10 Consignee vs Unit Price", pivot_consignee_price)
        r4 = write_table(ws3, "Top 10 Exporter vs Unit Price", pivot_exporter_price)
        r5 = write_table(ws3, "Top 10 Exporter vs Quantity", pivot_exporter_qty)
        r6 = write_table(ws3, "Classification vs Quantity", pivot_classification)

        # charts
        def add_chart(ws, start_row, title):
            chart = BarChart()
            chart.title = title

            data = Reference(ws, min_col=2, min_row=start_row+1, max_row=start_row+11)
            cats = Reference(ws, min_col=1, min_row=start_row+2, max_row=start_row+11)

            chart.add_data(data)
            chart.set_categories(cats)

            ws.add_chart(chart, f"E{start_row}")

        add_chart(ws3, r1, "Country vs Price")
        add_chart(ws3, r2, "Country vs Quantity")
        add_chart(ws3, r3, "Consignee vs Price")
        add_chart(ws3, r4, "Exporter vs Price")
        add_chart(ws3, r5, "Exporter vs Quantity")
        add_chart(ws3, r6, "Classification vs Quantity")

        # SAVE
        output = "final_output.xlsx"
        wb.save(output)

        st.success("✅ Fully Processed with Analysis & Charts")

        with open(output, "rb") as f:
            st.download_button(
                "⬇ Download Final Excel",
                f,
                file_name="final_dataset.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")