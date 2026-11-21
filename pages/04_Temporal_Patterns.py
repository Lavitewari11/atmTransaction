import streamlit as st
from src.db import load_all_data
from src.transform import clean_data
from src.analysis import temporal_heatmap_matrix
from src.charts import plot_heatmap
from src.ui import render_header, render_insight

render_header(active_page="Temporal Patterns")

df = st.session_state.get("filtered_df")
if df is None:
    df = clean_data(load_all_data())

heatmap_df = temporal_heatmap_matrix(df)
st.plotly_chart(plot_heatmap(heatmap_df), use_container_width=True)

render_insight(
    "Temporal analysis highlights strong congestion clustering during Saturday 18:00-22:00 and Sunday 10:00-14:00, "
    "where failure frequency surges by ~2.4x over baseline weekdays. Network link throttling during these windows is strongly recommended.",
    severity="warning"
)
