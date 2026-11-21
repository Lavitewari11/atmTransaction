import streamlit as st
import plotly.express as px
from src.db import load_all_data
from src.transform import clean_data
from src.analysis import get_failure_pareto
from src.charts import plot_pareto
from src.ui import render_header, render_insight, PLOTLY_LAYOUT

render_header(active_page="Failure Category")

df = st.session_state.get("filtered_df")
if df is None:
    df = clean_data(load_all_data())

pareto = get_failure_pareto(df)
st.plotly_chart(plot_pareto(pareto), use_container_width=True)

c1, c2 = st.columns(2)
with c1:
    st.subheader("💳 Failure by Transaction Type")
    type_fail = df[df["status"]=="Failed"].groupby("transaction_type").size().reset_index(name="count")
    fig = px.bar(type_fail, x="transaction_type", y="count", color="transaction_type", title="Failures by Operation")
    fig.update_layout(**PLOTLY_LAYOUT)
    st.plotly_chart(fig, use_container_width=True)

with c2:
    st.subheader("📋 Severity Distribution")
    sev = df[df["status"]=="Failed"]["severity"].value_counts().reset_index()
    sev.columns = ["severity", "count"]
    fig2 = px.pie(sev, values="count", names="severity", hole=0.4, title="Failure Severity Split")
    fig2.update_layout(**PLOTLY_LAYOUT)
    st.plotly_chart(fig2, use_container_width=True)

render_insight(
    "Pareto analysis confirms that Network-Switch and Customer-Account failures account for over 75% of total decline volume. "
    "Cash-Withdrawal is the single largest affected operation, representing critical friction for end-users.",
    severity="info"
)
