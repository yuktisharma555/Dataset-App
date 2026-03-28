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
        # ✅ Auto-detect correct header row
        df = None
        for i in range(6):
            temp_df = pd.read_excel(file, header=i)
            temp_df.columns = temp_df.columns.astype(str).str.strip().str.upper()

            if "GOODS DESCRIPTION" in temp_df.columns:
                df = temp_df
                break

        if df is None:
            st.error("❌ Could not detect correct header row")
            st.stop()

        # ✅ Clean dataset
        df = df.dropna(how="all")
        df.columns = df.columns.str.strip().str.upper()

        st.subheader("📌 Data Preview")
        st.dataframe(df.head())

        # ✅ Flexible column detection
        def find_col(possible_names):
            for col in df.columns:
                for name in possible_names:
                    if name in col:
                        return col
            return None

        desc_col = find_col(["DESCRIPTION"])
        qty_col = find_col(["QTY", "QUANTITY"])
        price_col = find_col(["PRICE"])
        curr_col = find_col(["CURR"])

        if not all([desc_col, qty_col, price_col, curr_col]):
            st.error("❌ Required columns not found properly")
            st.stop()

        # ✅ Clean numeric columns
        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce")
        df[price_col] = pd.to_numeric(df[price_col], errors="coerce")

        # ✅ Apply logic
        df["Classification"] = df[desc_col].apply(classify)
        df["Std Qty"] = df[qty_col].fillna(0)

        df["USD Price"] = df.apply(
            lambda x: x[price_col] * exchange_rates.get(str(x[curr_col]).strip(), 1),
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

        # ✅ FIXED DOWNLOAD BUTTONS

        # Excel
        with open(excel_file, "rb") as f:
            st.download_button(
                label="⬇ Download Processed Excel",
                data=f,
                file_name="processed_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # PPT
        with open(ppt_file, "rb") as f:
            st.download_button(
                label="⬇ Download PPT",
                data=f,
                file_name="analysis.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    except Exception as e:
        st.error(f"❌ Error occurred: {str(e)}")