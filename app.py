import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# ---------------- UI ----------------
st.set_page_config(page_title="Dataset Analyzer", layout="wide")

st.markdown("""
    <h1 style='text-align:center;color:#4CAF50;'>📊 Smart Dataset Cleaner</h1>
    <p style='text-align:center;'>Upload → Clean → Analyze → Download</p>
""", unsafe_allow_html=True)

file = st.file_uploader("📂 Upload Dataset", type=["xlsx"])
dataset_type = st.selectbox("📊 Dataset Type", ["medicine","vaccine","testkit"])

# ---------------- COLORS ----------------
dark_red = PatternFill("solid", fgColor="8B0000")
light_red = PatternFill("solid", fgColor="FF9999")
orange = PatternFill("solid", fgColor="FFA500")

exchange_rates = {
    "ZAR":0.0628,"THB":0.0322,"CAD":0.7334,"INR":0.0109,
    "MXN":0.058,"RUB":0.0129,"GBP":1.3455,"EUR":1.1821,
    "AED":0.2722,"USD":1
}

# ---------------- CLASSIFICATION ----------------
def classify(desc):
    d = str(desc).lower()

    if dataset_type == "medicine":
        if any(x in d for x in ["lamivudine","efavirenz","emtricitabine","dolutegravir","rilpivirine"]):
            return "FFP Combination"
        elif any(x in d for x in ["tablet","capsule","vial","bottle"]):
            return "FFP Plain"
        else:
            return "API"

    elif dataset_type == "vaccine":
        return "Pediatric Dose" if "pediatric" in d else "Adult Dose"

    else:
        if "strip" in d: return "Strip"
        if "card" in d: return "Card"
        return "Kit"

# ---------------- MAIN ----------------
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

        # column finder
        def find(keys):
            for c in df.columns:
                for k in keys:
                    if k in c:
                        return c
            return None

        desc = find(["DESCRIPTION"])
        qty = find(["QUANTITY","QTY"])
        price = find(["PRICE"])
        curr = find(["CURR"])
        unit = find(["UNIT"])
        exporter = find(["EXPORTER"])
        consignee = find(["CONSIGNEE"])
        country = find(["COUNTRY"])

        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        # NEW COLUMNS
        df["Classification"] = df[desc].apply(classify)
        df["USD Price"] = df.apply(lambda x: x[price]*exchange_rates.get(str(x[curr]).strip(),1), axis=1)

        df["Std Qty (NOS/PCS/VIAL)"] = df[qty]

        # ---------------- ANALYSIS ----------------
        def top10(col, val):
            return df.groupby(col)[val].sum().nlargest(10).reset_index() if col else None

        p1 = top10(country, "USD Price")
        p2 = top10(country, "Std Qty (NOS/PCS/VIAL)")
        p3 = top10(consignee, "USD Price")
        p4 = top10(exporter, "USD Price")
        p5 = top10(exporter, "Std Qty (NOS/PCS/VIAL)")
        p6 = df.groupby("Classification")["Std Qty (NOS/PCS/VIAL)"].sum().reset_index()

        # ---------------- EXCEL ----------------
        wb = Workbook()

        # ORIGINAL
        ws1 = wb.active
        ws1.title = "Original"
        for r in dataframe_to_rows(original, index=False, header=True):
            ws1.append(r)

        # CLEANED
        ws2 = wb.create_sheet("Cleaned")
        for r in dataframe_to_rows(df, index=False, header=True):
            ws2.append(r)

        # HEADER COLOR (NEW COLUMNS ONLY)
        for i, col in enumerate(df.columns, 1):
            if col in ["Classification","USD Price","Std Qty (NOS/PCS/VIAL)"]:
                ws2.cell(1,i).fill = orange
                ws2.cell(1,i).font = Font(bold=True)

        desc_index = list(df.columns).index(desc) + 1

        # ROW COLORING
        for i, row in df.iterrows():
            r = i + 2
            d = str(row[desc]).lower()
            q = row[qty]
            p = row[price]

            cell = ws2.cell(r, desc_index)

            if pd.notna(q) and q < 1:
                cell.fill = dark_red

            if pd.notna(p) and p == 0:
                cell.fill = orange

            if any(x in d for x in ["sample","standard","r&d","impurity"]):
                cell.fill = dark_red

            if dataset_type == "medicine" and "api" in str(row["Classification"]).lower():
                cell.fill = light_red

            # 🔥 FIXED VACCINE RULE
            if dataset_type == "vaccine" and ("free of cost" in d or "foc" in d):
                cell.fill = dark_red

        # ANALYSIS
        ws3 = wb.create_sheet("Analysis")

        def write(ws, title, data):
            ws.append([title])
            start = ws.max_row

            if data is not None:
                for r in dataframe_to_rows(data, index=False, header=True):
                    ws.append(r)
            else:
                ws.append(["Not Available"])

            ws.append([])
            return start

        r1 = write(ws3,"Country vs Price",p1)
        r2 = write(ws3,"Country vs Qty",p2)
        r3 = write(ws3,"Consignee vs Price",p3)
        r4 = write(ws3,"Exporter vs Price",p4)
        r5 = write(ws3,"Exporter vs Qty",p5)
        r6 = write(ws3,"Classification vs Qty",p6)

        def chart(ws, row, title):
            c = BarChart()
            c.title = title
            data = Reference(ws, min_col=2, min_row=row+1, max_row=row+11)
            cats = Reference(ws, min_col=1, min_row=row+2, max_row=row+11)
            c.add_data(data)
            c.set_categories(cats)
            ws.add_chart(c, f"E{row}")

        chart(ws3,r1,"Country Price")
        chart(ws3,r2,"Country Qty")
        chart(ws3,r3,"Consignee Price")
        chart(ws3,r4,"Exporter Price")
        chart(ws3,r5,"Exporter Qty")
        chart(ws3,r6,"Classification Qty")

        wb.save("final.xlsx")

        st.success("✅ Done Successfully")

        with open("final.xlsx","rb") as f:
            st.download_button("⬇ Download Excel",f,"final.xlsx")

    except Exception as e:
        st.error(str(e))