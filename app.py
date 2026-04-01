import streamlit as st
import pandas as pd
import plotly.express as px

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Dataset Dashboard", layout="wide")

# ---------------- PREMIUM UI (GLASSMORPHISM) ----------------
st.markdown("""
<style>

/* Background */
body {
    background: linear-gradient(135deg, #e0f2fe, #ffffff, #ede9fe);
}

/* Title */
.title {
    font-size:40px;
    font-weight:700;
    text-align:center;
    color:#1e293b;
    margin-bottom:20px;
}

/* Glass Cards */
.card {
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    background: rgba(255, 255, 255, 0.25);
    border-radius: 16px;
    padding: 20px;
    border: 1px solid rgba(255,255,255,0.3);
    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
    text-align:center;
}

/* KPI text */
.kpi {
    font-size:28px;
    font-weight:bold;
    color:#0f172a;
}

/* Labels */
.label {
    font-size:14px;
    color:#475569;
}

</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>📊 Smart Dataset Dashboard</div>", unsafe_allow_html=True)

# ---------------- FILE UPLOAD ----------------
file = st.file_uploader("Upload Excel File", type=["xlsx"])

# ---------------- MAIN ----------------
if file:
    try:
        # Load data (header row = 6)
        df = pd.read_excel(file, header=5)
        df.columns = df.columns.str.upper().str.strip()

        # Required columns
        desc = "GOODS DESCRIPTION"
        qty = "QUANTITY"
        price = "TOTAL PRICE_INV_FC"
        country = "COUNTRY"
        exporter = "EXPORTER"
        consignee = "CONSIGNEE NAME"

        # Convert numeric
        df[qty] = pd.to_numeric(df[qty], errors="coerce")
        df[price] = pd.to_numeric(df[price], errors="coerce")

        # ---------------- KPIs ----------------
        total_value = df[price].sum()
        total_qty = df[qty].sum()
        total_records = len(df)

        col1, col2, col3 = st.columns(3)

        col1.markdown(f"""
        <div class='card'>
            <div class='kpi'>{total_value:,.0f}</div>
            <div class='label'>Total Value</div>
        </div>
        """, unsafe_allow_html=True)

        col2.markdown(f"""
        <div class='card'>
            <div class='kpi'>{total_qty:,.0f}</div>
            <div class='label'>Total Quantity</div>
        </div>
        """, unsafe_allow_html=True)

        col3.markdown(f"""
        <div class='card'>
            <div class='kpi'>{total_records}</div>
            <div class='label'>Total Records</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        # ---------------- FILTERS ----------------
        st.subheader("Filters")

        col1, col2, col3 = st.columns(3)

        selected_country = col1.multiselect("Country", df[country].dropna().unique())
        selected_exporter = col2.multiselect("Exporter", df[exporter].dropna().unique())
        selected_consignee = col3.multiselect("Consignee", df[consignee].dropna().unique())

        filtered_df = df.copy()

        if selected_country:
            filtered_df = filtered_df[filtered_df[country].isin(selected_country)]

        if selected_exporter:
            filtered_df = filtered_df[filtered_df[exporter].isin(selected_exporter)]

        if selected_consignee:
            filtered_df = filtered_df[filtered_df[consignee].isin(selected_consignee)]

        # ---------------- CHARTS ----------------
        st.markdown("## Analysis")

        col1, col2 = st.columns(2)

        # Country vs Value
        c1 = filtered_df.groupby(country)[price].sum().nlargest(10).reset_index()
        fig1 = px.bar(
            c1, x=country, y=price,
            title="Top 10 Countries by Value",
            color=price
        )
        col1.plotly_chart(fig1, use_container_width=True)

        # Country vs Quantity
        c2 = filtered_df.groupby(country)[qty].sum().nlargest(10).reset_index()
        fig2 = px.bar(
            c2, x=country, y=qty,
            title="Top 10 Countries by Quantity",
            color=qty
        )
        col2.plotly_chart(fig2, use_container_width=True)

        col1, col2 = st.columns(2)

        # Exporter vs Value
        c3 = filtered_df.groupby(exporter)[price].sum().nlargest(10).reset_index()
        fig3 = px.bar(
            c3, x=exporter, y=price,
            title="Top Exporters by Value",
            color=price
        )
        col1.plotly_chart(fig3, use_container_width=True)

        # Consignee vs Value
        c4 = filtered_df.groupby(consignee)[price].sum().nlargest(10).reset_index()
        fig4 = px.bar(
            c4, x=consignee, y=price,
            title="Top Consignees by Value",
            color=price
        )
        col2.plotly_chart(fig4, use_container_width=True)

        st.markdown("---")

        # ---------------- DATA TABLE ----------------
        st.subheader("Filtered Dataset")
        st.dataframe(filtered_df, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")