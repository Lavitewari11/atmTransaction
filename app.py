import streamlit as st
import pandas as pd
from src.db import load_all_data
from src.transform import audit_data_quality, clean_data
from src.ui import render_header, render_sidebar_filters, apply_filters, quality_warning

st.set_page_config(
    page_title="ATM Transaction Failure Analysis",
    page_icon="🏧",
    layout="wide",
    initial_sidebar_state="expanded"
)

render_header(active_page="Fleet Overview")

df_raw = load_all_data()

quality = audit_data_quality(df_raw)
st.sidebar.markdown("### 🛡️ Data Quality Telemetry")
quality_warning(f"Missing Latencies: {quality['missing_latency']:,}")
quality_warning(f"Quarantined Duplicates: {quality['duplicate_txns']:,}")
quality_warning(f"Extreme Outlier Amounts: {quality['outlier_amounts']:,}")

df_clean = clean_data(df_raw)
filters = render_sidebar_filters(df_clean)
df_filtered = apply_filters(df_clean, filters)
st.session_state["filtered_df"] = df_filtered

st.info("👈 Use the left sidebar to navigate across analytical modules or slice filters!")
st.markdown("### 📊 Active Filter Summary")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Analyzed Records", f"{len(df_filtered):,}")
c2.metric("Date Span", f"{filters['date_range'][0]} to {filters['date_range'][1]}" if len(filters['date_range'])==2 else "All")
c3.metric("Selected Banks", f"{len(filters['banks'])} Banks")
c4.metric("Selected Zones", f"{len(filters['zones'])} Zones")

st.markdown("---")
st.subheader("📥 Export Filtered Audit Subsets")
csv = df_filtered.to_csv(index=False).encode('utf-8')
st.download_button(
    label="⬇️ Download Filtered Transactions (CSV)",
    data=csv,
    file_name="filtered_atm_transactions.csv",
    mime="text/csv"
)
