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
        # ---- ORIGINAL DATA ----
        original_df = pd.read_excel(file)

        # ---- CLEAN DATA ----
        df = None
        for i in range(6):
            temp_df = pd.read_excel(file, header=i)
            temp_df.columns = temp_df.columns.astype(str).str.strip().str.upper()
            if "GOODS DESCRIPTION" in temp_df.columns:
                df = temp_df
                break

        if df is None:
            st.error("Header not found")
            st.stop()

        df = df.dropna(how="all")

        # Detect columns
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

        # Convert numeric
        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        # ---- CLEANING RULES ----

        # Classification
        df["Classification"] = df[desc].apply(classify)

        # USD Price
        df["USD Price"] = df.apply(
            lambda x: x[price] * exchange_rates.get(str(x[curr]).strip(),1),
            axis=1
        )

        # Flags (basic version of your rules)
        df["Flag"] = ""

        df.loc[df[price] == 0, "Flag"] = "Zero Value"
        df.loc[df[desc].str.contains("sample|standard|r&d", case=False, na=False), "Flag"] = "Special Case"
        df.loc[df[desc].str.contains("impurity", case=False, na=False), "Flag"] = "Impurity"

        # ---- ANALYSIS ----
        analysis = df.groupby("Classification").agg(
            Total_Records=("Classification","count"),
            Total_Value=("USD Price","sum")
        ).reset_index()

        # ---- SAVE MULTI-SHEET EXCEL ----
        output_file = "final_output.xlsx"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            original_df.to_excel(writer, sheet_name="Original Data", index=False)
            df.to_excel(writer, sheet_name="Cleaned Data", index=False)
            analysis.to_excel(writer, sheet_name="Analysis", index=False)

        # ---- UI ----
        st.success("✅ Full Processing Done")

        st.subheader("Cleaned Data Preview")
        st.dataframe(df.head())

        st.subheader("Analysis Preview")
        st.dataframe(analysis)

        # ---- DOWNLOAD ----
        with open(output_file, "rb") as f:
            st.download_button(
                "⬇ Download Full Excel",
                f,
                file_name="dataset_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {str(e)}")