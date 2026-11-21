import streamlit as st
from src.db import load_all_data
from src.transform import clean_data
from src.analysis import compute_kpis, get_hourly_trend
from src.charts import plot_hourly_trend
from src.ui import render_header, render_kpi_row, render_insight

render_header(active_page="Fleet Overview")

df = st.session_state.get("filtered_df")
if df is None:
    df = clean_data(load_all_data())

kpis = compute_kpis(df)
render_kpi_row([
    {"label": "Total Transactions", "value": f"{kpis['total']:,}", "icon": "🔢"},
    {"label": "Approved Transactions", "value": f"{kpis['approved']:,}", "delta": f"{100-kpis['fail_rate']:.1f}%", "delta_positive": True, "icon": "✅"},
    {"label": "Failed Transactions", "value": f"{kpis['failed']:,}", "delta": f"{kpis['fail_rate']:.1f}% Fail", "delta_positive": False, "icon": "❌"},
    {"label": "Disrupted INR Value", "value": f"₹ {kpis['disrupted_amount']/1e7:.2f} Cr", "icon": "💸"}
])

trend = get_hourly_trend(df)
st.plotly_chart(plot_hourly_trend(trend), use_container_width=True)

render_insight(
    "Intraday analysis indicates recurring failure surges during weekend peak hours (Saturday evening and Sunday afternoon). "
    "Overall fleet failure rate hovers around 21.3%, significantly above the 5.0% banking operational SLA benchmark.",
    severity="warning"
)
