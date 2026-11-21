# src/ui.py
# p096: ATM Transaction Failure Analysis
# Reusable UI components — Tailwind header + KPI cards via st.components.v1.html

import streamlit as st
import streamlit.components.v1 as components

# ─────────────────────────── THEME ───────────────────────────────────────────
THEME = {
    "primary":    "#2B6CB0",
    "danger":     "#E53E3E",
    "success":    "#38A169",
    "warning":    "#D69E2E",
    "bg":         "#F7FAFC",
    "sidebar_bg": "#EBF4FF",
    "text":       "#1A202C",
    "subtext":    "#4A5568",
}
PLOTLY_COLORS = ["#2B6CB0", "#E53E3E", "#D69E2E", "#38A169", "#805AD5", "#DD6B20"]

PLOTLY_LAYOUT = dict(
    plot_bgcolor="#F7FAFC",
    paper_bgcolor="#FFFFFF",
    font=dict(family="Inter, Calibri, sans-serif", color="#1A202C"),
    colorway=PLOTLY_COLORS,
    margin=dict(l=40, r=20, t=40, b=40),
    legend=dict(bgcolor="rgba(0,0,0,0)", borderwidth=0),
    hoverlabel=dict(bgcolor="#1A365D", font_color="white", font_family="Inter"),
)

PLOTLY_CONFIG = {
    "displayModeBar": True,
    "toImageButtonOptions": {"format": "svg", "filename": "p096_chart"},
    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
}

# ─────────────────────────── TAILWIND HEADER ─────────────────────────────────
def render_header(active_page: str = "Fleet Overview") -> None:
    """Render the full-width Tailwind top navigation bar."""
    pages = [
        ("Fleet Overview",    "01_Fleet_Overview"),
        ("Failure Category",  "02_Failure_Category"),
        ("Terminal Health",   "03_Terminal_Health"),
        ("Temporal Patterns", "04_Temporal_Patterns"),
        ("Network & Latency", "05_Network_Latency"),
        ("Financial Impact",  "06_Financial_Impact"),
    ]
    nav_items = ""
    for label, _ in pages:
        active_cls = (
            "bg-blue-800 text-white font-semibold"
            if label == active_page
            else "text-blue-100 hover:bg-blue-700 hover:text-white"
        )
        nav_items += (
            f'<span class="{active_cls} px-3 py-1 rounded-md text-sm cursor-pointer">'
            f"{label}</span>"
        )

    html = f"""
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
    <div class="w-full bg-blue-900 text-white px-6 py-3 rounded-xl shadow-lg mb-4">
      <div class="flex items-center justify-between flex-wrap gap-2">
        <div class="flex items-center gap-3">
          <span class="text-2xl">🏧</span>
          <div>
            <p class="font-bold text-lg leading-tight tracking-wide">ATM Transaction Failure Analysis</p>
            <p class="text-blue-300 text-xs">p096 · Lavi Tewari · ABES MCA 2nd Yr · Mentor: Mr. Dinesh Sereno</p>
          </div>
        </div>
        <div class="flex flex-wrap gap-1">{nav_items}</div>
      </div>
    </div>
    """
    components.html(html, height=100, scrolling=False)

# ─────────────────────────── KPI CARD ────────────────────────────────────────
def kpi_card(label: str, value: str, delta: str = "", delta_positive: bool = True,
             icon: str = "📊", bg_color: str = "#FFFFFF") -> str:
    """Return HTML for a single Tailwind KPI card."""
    delta_color = "text-green-600" if delta_positive else "text-red-600"
    delta_arrow = "▲" if delta_positive else "▼"
    delta_html = (
        f'<p class="text-xs {delta_color} mt-1 font-medium">{delta_arrow} {delta}</p>'
        if delta else ""
    )
    return f"""
    <div class="bg-white rounded-xl shadow-md p-4 flex flex-col border-l-4 border-blue-600 flex-1"
         style="background-color:{bg_color}; min-width:140px;">
      <div class="flex items-center justify-between mb-1">
        <span class="text-2xl">{icon}</span>
        <span class="text-xs text-gray-400 uppercase tracking-wider font-semibold">{label}</span>
      </div>
      <p class="text-2xl font-bold text-gray-800 leading-tight">{value}</p>
      {delta_html}
    </div>
    """

def render_kpi_row(cards: list) -> None:
    """Render a horizontal row of KPI cards via Tailwind HTML."""
    cards_html = "".join(kpi_card(**c) for c in cards)
    html = f"""
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
    <div class="flex flex-wrap gap-4 mb-4">
      {cards_html}
    </div>
    """
    components.html(html, height=125, scrolling=False)

# ─────────────────────────── INSIGHT BOX ─────────────────────────────────────
def render_insight(text: str, severity: str = "info") -> None:
    """Render a Tailwind-styled insight/callout box below charts."""
    colors = {
        "info":    ("border-blue-500",  "bg-blue-50",  "text-blue-800",  "💡"),
        "warning": ("border-yellow-500","bg-yellow-50","text-yellow-800","⚠️"),
        "danger":  ("border-red-500",   "bg-red-50",   "text-red-800",   "🚨"),
        "success": ("border-green-500", "bg-green-50", "text-green-800", "✅"),
    }
    border, bg, txt, icon = colors.get(severity, colors["info"])
    html = f"""
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
    <div class="{bg} border-l-4 {border} rounded-r-xl px-4 py-3 mb-3 shadow-sm">
      <p class="{txt} text-sm font-medium">{icon}  {text}</p>
    </div>
    """
    components.html(html, height=65, scrolling=False)

# ─────────────────────────── OUTLIER BADGE ───────────────────────────────────
def outlier_badge(terminal_id: str, failure_rate: float) -> str:
    """Return an inline HTML red pill badge for a chronic outlier terminal."""
    return (
        f'<span style="background:#FED7D7;color:#9B2C2C;border-radius:9999px;'
        f'padding:2px 10px;font-size:11px;font-weight:700;font-family:Inter,sans-serif;">'
        f"⚠ {terminal_id} ({failure_rate:.1f}%)</span>"
    )

# ─────────────────────────── QUALITY BADGE ───────────────────────────────────  
def quality_warning(message: str) -> None:
    """Show a data quality amber warning badge in the sidebar."""
    st.sidebar.markdown(
        f'<div style="background:#FEFCBF;border-left:4px solid #D69E2E;'
        f'border-radius:4px;padding:6px 10px;font-size:12px;color:#744210;margin-bottom:10px;">'
        f"⚠️ {message}</div>",
        unsafe_allow_html=True,
    )

# ─────────────────────────── SIDEBAR FILTERS ─────────────────────────────────
def render_sidebar_filters(df) -> dict:
    """Render all global sidebar filters and return a dict of selected values."""
    import pandas as pd
    st.sidebar.markdown("## 🔍 Filters")
    st.sidebar.markdown("---")

    min_date = df["timestamp"].min().date()
    max_date = df["timestamp"].max().date()
    date_range = st.sidebar.date_input(
        "📅 Date Range", value=(min_date, max_date),
        min_value=min_date, max_value=max_date
    )

    all_banks = sorted(df["bank_owner"].dropna().unique().tolist())
    banks = st.sidebar.multiselect("🏦 Bank", options=all_banks, default=all_banks)

    all_zones = sorted(df["zone"].dropna().unique().tolist())
    zones = st.sidebar.multiselect("🗺️ Zone", options=all_zones, default=all_zones)

    all_types = sorted(df["transaction_type"].dropna().unique().tolist())
    tx_types = st.sidebar.multiselect("💳 Transaction Type", options=all_types, default=all_types)

    all_cats = sorted(df["failure_category"].dropna().unique().tolist())
    failure_cats = st.sidebar.multiselect(
        "⚠️ Failure Category", options=all_cats, default=all_cats
    )

    card_types = st.sidebar.multiselect(
        "🔄 Card Network", options=["On-Us","Off-Us"], default=["On-Us","Off-Us"]
    )

    if st.sidebar.button("🔄 Reset All Filters"):
        st.rerun()

    return {
        "date_range": date_range,
        "banks": banks,
        "zones": zones,
        "tx_types": tx_types,
        "failure_cats": failure_cats,
        "card_types": card_types,
    }

def apply_filters(df, filters: dict):
    """Apply sidebar filter dict to the full DataFrame and return filtered copy."""
    fdf = df.copy()
    if len(filters["date_range"]) == 2:
        start, end = filters["date_range"]
        fdf = fdf[(fdf["timestamp"].dt.date >= start) & (fdf["timestamp"].dt.date <= end)]
    if filters["banks"]:
        fdf = fdf[fdf["bank_owner"].isin(filters["banks"])]
    if filters["zones"]:
        fdf = fdf[fdf["zone"].isin(filters["zones"])]
    if filters["tx_types"]:
        fdf = fdf[fdf["transaction_type"].isin(filters["tx_types"])]
    if filters["failure_cats"]:
        failed = fdf["status"] == "Failed"
        in_cat = fdf["failure_category"].isin(filters["failure_cats"])
        fdf = fdf[~failed | (failed & in_cat)]
    if filters["card_types"]:
        fdf = fdf[fdf["card_type"].isin(filters["card_types"])]
    return fdf
