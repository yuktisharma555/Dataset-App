import streamlit as st
import pandas as pd
from pptx import Presentation

st.set_page_config(page_title="Dataset Cleaner", layout="wide")

st.title("📊 Dataset Cleaner & Analyzer")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

dataset_type = st.selectbox(
    "Select Dataset Type",
    ["medicine", "vaccine", "testkit"]
)

exchange_rates = {
    "ZAR": 0.0628, "THB": 0.0322, "CAD": 0.7334,
    "INR": 0.0109, "MXN": 0.058, "RUB": 0.0129,
    "GBP": 1.3455, "EUR": 1.1821, "AED": 0.2722, "USD": 1
}

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

if file:
    try:
        # Try multiple header positions
        for i in range(6):
            df = pd.read_excel(file, header=i)
            df.columns = df.columns.astype(str).str.strip().str.upper()

            if "GOODS DESCRIPTION" in df.columns:
                break

        df = df.dropna(how="all")

        st.subheader("📌 Data Preview")
        st.dataframe(df.head())

        # Flexible column detection
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

        # Clean data
        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce")
        df[price_col] = pd.to_numeric(df[price_col], errors="coerce")

        # Processing
        df["Classification"] = df[desc_col].apply(classify)
        df["Std Qty"] = df[qty_col].fillna(0)

        df["USD Price"] = df.apply(
            lambda x: x[price_col] * exchange_rates.get(str(x[curr_col]).strip(), 1),
            axis=1
        )

        total_records = len(df)
        total_value = df["USD Price"].sum()

        st.success("✅ Processing Complete")

        st.write(f"Total Records: {total_records}")
        st.write(f"Total USD Value: {round(total_value,2)}")

        # Save Excel
        excel_file = "processed_output.xlsx"
        df.to_excel(excel_file, index=False)

        # PPT
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Dataset Analysis"
        slide.placeholders[1].text = f"Records: {total_records}\nValue: {round(total_value,2)}"

        ppt_file = "analysis.pptx"
        prs.save(ppt_file)

        with open(excel_file, "rb") as f:
            st.download_button("⬇ Download Excel", f)

        with open(ppt_file, "rb") as f:
            st.download_button("⬇ Download PPT", f)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")