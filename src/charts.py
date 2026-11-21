import plotly.express as px
import plotly.graph_objects as go
from src.ui import PLOTLY_LAYOUT

def plot_hourly_trend(trend_df):
    fig = px.line(
        trend_df, x="hour_dt", y="failure_rate",
        title="Fleet Hourly Transaction Failure Rate (%)",
        labels={"hour_dt": "Timeline", "failure_rate": "Failure Rate (%)"}
    )
    fig.update_traces(line=dict(color="#E53E3E", width=2))
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def plot_pareto(pareto_df):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=pareto_df["failure_category"], y=pareto_df["count"],
        name="Failure Volume", marker_color="#2B6CB0"
    ))
    fig.add_trace(go.Scatter(
        x=pareto_df["failure_category"], y=pareto_df["cumulative_pct"],
        name="Cumulative %", yaxis="y2", mode="lines+markers",
        line=dict(color="#E53E3E", width=3)
    ))
    layout = PLOTLY_LAYOUT.copy()
    layout.update(dict(
        title="Failure Root Causes (Pareto 80/20 Analysis)",
        yaxis=dict(title="Volume"),
        yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105])
    ))
    fig.update_layout(**layout)
    return fig

def plot_heatmap(heatmap_df):
    fig = px.imshow(
        heatmap_df,
        labels=dict(x="Hour of Day (0-23)", y="Day of Week", color="Failures"),
        color_continuous_scale="Reds",
        title="24x7 Temporal Failure Congestion Heatmap"
    )
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def plot_latency_scatter(df, sample_n=2000):
    sample = df.dropna(subset=["response_time_ms"]).sample(n=min(len(df), sample_n), random_state=42)
    fig = px.scatter(
        sample, x="response_time_ms", y="status", color="status",
        color_discrete_map={"Approved": "#38A169", "Failed": "#E53E3E"},
        title="Network Response Time (ms) vs. Transaction Outcome",
        labels={"response_time_ms": "Response Time (ms)", "status": "Outcome"}
    )
    fig.add_vline(x=3500, line_dash="dash", line_color="#D69E2E", annotation_text="Timeout (3.5s)")
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def plot_onus_offus_donut(df):
    grouped = df.groupby(["card_type", "status"]).size().unstack(fill_value=0)
    fig = px.pie(
        values=grouped["Failed"], names=grouped.index, hole=0.4,
        title="Failure Volume Split: On-Us vs. Off-Us",
        color_discrete_sequence=["#2B6CB0", "#DD6B20"]
    )
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def plot_financial_treemap(df):
    failed = df[df["status"] == "Failed"]
    fig = px.treemap(
        failed, path=["failure_category", "transaction_type"], values="amount",
        color="failure_category",
        title="Financial Value Disrupted (INR) by Category & Operation"
    )
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig
