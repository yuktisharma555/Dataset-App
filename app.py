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

# COLORS
dark_red = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
light_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

def classify(desc, price):
    desc = str(desc).lower()

    if dataset_type == "medicine":
        if any(x in desc for x in ["lamivudine","efavirenz","emtricitabine","dolutegravir","rilpivirine"]):
            return "FFP Combination"

        if any(x in desc for x in ["tablet","capsule","vial","bottle"]):
            return "FFP Plain"

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
        original = pd.read_excel(file)

        # detect header
        for i in range(6):
            df = pd.read_excel(file, header=i)
            df.columns = df.columns.str.upper().str.strip()
            if "GOODS DESCRIPTION" in df.columns:
                break

        df = df.dropna(how="all")

        # detect columns
        def find(name):
            return [c for c in df.columns if name in c][0]

        desc = find("DESCRIPTION")
        qty = find("QUANTITY")
        price = find("PRICE")
        curr = find("CURR")

        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        # new columns
        df["Classification"] = df.apply(lambda x: classify(x[desc], x[price]), axis=1)
        df["USD Price"] = df.apply(lambda x: x[price]*exchange_rates.get(str(x[curr]).strip(),1), axis=1)
        df["Std Qty"] = df[qty]

        # ANALYSIS
        top_exporter = df.groupby("EXPORTER NAME")["USD Price"].sum().nlargest(5)
        top_consignee = df.groupby("CONSIGNEE NAME")["USD Price"].sum().nlargest(5)
        top_country = df.groupby("COUNTRY")["USD Price"].sum().nlargest(5)

        # CREATE EXCEL
        wb = Workbook()

        # ---- ORIGINAL ----
        ws1 = wb.active
        ws1.title = "Original Data"

        for r in dataframe_to_rows(original, index=False, header=True):
            ws1.append(r)

        # ---- CLEANED ----
        ws2 = wb.create_sheet("Cleaned Data")

        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # APPLY COLORS
        for i, row in enumerate(df.itertuples(), start=2):

            desc_val = str(getattr(row, desc.replace(" ","_"))).lower()
            qty_val = getattr(row, qty.replace(" ","_"))
            price_val = getattr(row, price.replace(" ","_"))

            cell = ws2[f"C{i}"]

            if qty_val and qty_val < 1:
                cell.fill = dark_red

            if price_val == 0:
                cell.fill = orange

            if any(x in desc_val for x in ["sample","standard","r&d","impurity"]):
                cell.fill = dark_red

            if dataset_type == "medicine" and "api" in row.Classification.lower():
                cell.fill = light_red

        # ---- ANALYSIS ----
        ws3 = wb.create_sheet("Analysis")

        ws3.append(["Top Exporters"])
        for k,v in top_exporter.items():
            ws3.append([k,v])

        ws3.append([])
        ws3.append(["Top Consignees"])
        for k,v in top_consignee.items():
            ws3.append([k,v])

        ws3.append([])
        ws3.append(["Top Countries"])
        for k,v in top_country.items():
            ws3.append([k,v])

        # SAVE
        output = "final_output.xlsx"
        wb.save(output)

        st.success("✅ Full Methodology Applied")

        with open(output, "rb") as f:
            st.download_button(
                "⬇ Download Final Excel",
                f,
                file_name="final_dataset.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(str(e))