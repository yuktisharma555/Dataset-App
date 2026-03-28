import streamlit as st
import pandas as pd
from pptx import Presentation

st.set_page_config(page_title="Dataset Cleaner", layout="wide")

st.title("📊 Dataset Cleaner & Analyzer")

# Upload file
file = st.file_uploader("Upload Excel File", type=["xlsx"])

dataset_type = st.selectbox(
    "Select Dataset Type",
    ["medicine", "vaccine", "testkit"]
)

# Exchange rates
exchange_rates = {
    "ZAR": 0.0628, "THB": 0.0322, "CAD": 0.7334,
    "INR": 0.0109, "MXN": 0.058, "RUB": 0.0129,
    "GBP": 1.3455, "EUR": 1.1821, "AED": 0.2722, "USD": 1
}

# Classification logic
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

    elif dataset_type == "testkit":
        if "strip" in desc:
            return "Strip"
        elif "card" in desc:
            return "Card"
        else:
            return "Kit"

# MAIN PROCESS
if file:
    try:
        # ✅ FIX 1: Correct header row
        df = pd.read_excel(file, header=4)

        # ✅ FIX 2: Clean dataset
        df = df.dropna(how="all")
        df.columns = df.columns.str.strip()

        st.subheader("📌 Cleaned Data Preview")
        st.dataframe(df.head())

        # ✅ Ensure required columns exist
        required_cols = ["GOODS DESCRIPTION", "QUANTITY", "ITEM PRICE_INV", "CURRENCY"]

        for col in required_cols:
            if col not in df.columns:
                st.error(f"❌ Missing column: {col}")
                st.stop()

        # ✅ FIX 3: Classification
        df["Classification"] = df["GOODS DESCRIPTION"].apply(classify)

        # ✅ FIX 4: Standard quantity
        df["Std Qty"] = pd.to_numeric(df["QUANTITY"], errors="coerce").fillna(0)

        # ✅ FIX 5: Numeric conversion
        df["ITEM PRICE_INV"] = pd.to_numeric(df["ITEM PRICE_INV"], errors="coerce")

        # ✅ FIX 6: USD conversion
        df["USD Price"] = df.apply(
            lambda x: x["ITEM PRICE_INV"] * exchange_rates.get(str(x["CURRENCY"]).strip(), 1),
            axis=1
        )

        # ✅ Analysis
        total_records = len(df)
        total_value = df["USD Price"].sum()

        st.success("✅ Processing Complete")

        st.subheader("📊 Analysis Summary")
        st.write(f"Total Records: {total_records}")
        st.write(f"Total USD Value: {round(total_value,2)}")

        # ✅ Save Excel
        excel_file = "processed_output.xlsx"
        df.to_excel(excel_file, index=False)

        # ✅ Create PPT
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        slide.shapes.title.text = "Dataset Analysis"
        slide.placeholders[1].text = f"""
Total Records: {total_records}
Total USD Value: {round(total_value,2)}
"""

        ppt_file = "analysis.pptx"
        prs.save(ppt_file)

        # ✅ Download buttons
        with open(excel_file, "rb") as f:
            st.download_button("⬇ Download Processed Excel", f, file_name="processed.xlsx")

        with open(ppt_file, "rb") as f:
            st.download_button("⬇ Download PPT", f, file_name="analysis.pptx")

    except Exception as e:
        st.error(f"❌ Error occurred: {str(e)}")