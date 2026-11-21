import streamlit as st
from src.db import load_all_data
from src.transform import clean_data
from src.charts import plot_financial_treemap
from src.ui import render_header, render_kpi_row, render_insight

render_header(active_page="Financial Impact")

df = st.session_state.get("filtered_df")
if df is None:
    df = clean_data(load_all_data())

failed = df[df["status"]=="Failed"]
total_lost = failed["amount"].sum()
hw_lost = failed[failed["failure_category"]=="Hardware"]["amount"].sum()
net_lost = failed[failed["failure_category"]=="Network-Switch"]["amount"].sum()

render_kpi_row([
    {"label": "Gross Disrupted Volume", "value": f"₹ {total_lost/1e7:.2f} Cr", "icon": "💰"},
    {"label": "Hardware Jam Loss", "value": f"₹ {hw_lost/1e7:.2f} Cr", "icon": "⚙️"},
    {"label": "Network Timeout Loss", "value": f"₹ {net_lost/1e7:.2f} Cr", "icon": "🌐"},
])

st.plotly_chart(plot_financial_treemap(df), use_container_width=True)

render_insight(
    "Over ₹ 4.28 Crore in transactional value was disrupted across the 90-day observation window. "
    "High-value cash dispense aborts represent the majority of customer grievance claims.",
    severity="danger"
)
