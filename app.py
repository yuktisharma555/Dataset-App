import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

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
        # ---- ORIGINAL DATA ----
        original_df = pd.read_excel(file)

        # ---- FIND CORRECT HEADER ----
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

        # ---- FIND COLUMNS FLEXIBLY ----
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
        exporter_col = find_col(["EXPORTER"])
        consignee_col = find_col(["CONSIGNEE"])
        country_col = find_col(["COUNTRY"])

        if not all([desc, qty, price, curr]):
            st.error("❌ Required columns not found")
            st.stop()

        # ---- CLEAN DATA ----
        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        df["Classification"] = df[desc].apply(classify)
        df["USD Price"] = df.apply(
            lambda x: x[price] * exchange_rates.get(str(x[curr]).strip(),1),
            axis=1
        )
        df["Std Qty"] = df[qty]

        # ---- ANALYSIS ----
        top_exporter = df.groupby(exporter_col)["USD Price"].sum().nlargest(5) if exporter_col else None
        top_consignee = df.groupby(consignee_col)["USD Price"].sum().nlargest(5) if consignee_col else None
        top_country = df.groupby(country_col)["USD Price"].sum().nlargest(5) if country_col else None

        # ---- CREATE EXCEL ----
        wb = Workbook()

        # Sheet 1: Original
        ws1 = wb.active
        ws1.title = "Original Data"
        for r in dataframe_to_rows(original_df, index=False, header=True):
            ws1.append(r)

        # Sheet 2: Cleaned
        ws2 = wb.create_sheet("Cleaned Data")
        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # ---- APPLY COLORING ----
        for idx, row in df.iterrows():
            desc_val = str(row[desc]).lower()
            qty_val = row[qty]
            price_val = row[price]

            excel_row = idx + 2  # header offset
            cell = ws2[f"C{excel_row}"]  # GOODS DESCRIPTION column assumed C

            if pd.notna(qty_val) and qty_val < 1:
                cell.fill = dark_red

            if pd.notna(price_val) and price_val == 0:
                cell.fill = orange

            if any(x in desc_val for x in ["sample","standard","r&d","impurity"]):
                cell.fill = dark_red

            if dataset_type == "medicine" and "api" in str(row["Classification"]).lower():
                cell.fill = light_red

        # Sheet 3: Analysis
        ws3 = wb.create_sheet("Analysis")

        ws3.append(["Top Exporters"])
        if top_exporter is not None:
            for k,v in top_exporter.items():
                ws3.append([k,v])
        else:
            ws3.append(["Not Available"])

        ws3.append([])
        ws3.append(["Top Consignees"])
        if top_consignee is not None:
            for k,v in top_consignee.items():
                ws3.append([k,v])
        else:
            ws3.append(["Not Available"])

        ws3.append([])
        ws3.append(["Top Countries"])
        if top_country is not None:
            for k,v in top_country.items():
                ws3.append([k,v])
        else:
            ws3.append(["Not Available"])

        # SAVE FILE
        output = "final_output.xlsx"
        wb.save(output)

        st.success("✅ Full Processing Done")

        # DOWNLOAD
        with open(output, "rb") as f:
            st.download_button(
                "⬇ Download Final Excel",
                f,
                file_name="final_dataset.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")