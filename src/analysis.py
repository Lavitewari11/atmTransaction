import pandas as pd
import numpy as np
from scipy import stats

def compute_kpis(df: pd.DataFrame) -> dict:
    total = len(df)
    failed = (df["status"] == "Failed").sum()
    approved = (df["status"] == "Approved").sum()
    fail_rate = (failed / total * 100) if total > 0 else 0.0
    disrupted_amount = df[df["status"] == "Failed"]["amount"].sum()
    return {
        "total": total,
        "failed": failed,
        "approved": approved,
        "fail_rate": fail_rate,
        "disrupted_amount": disrupted_amount
    }

def get_hourly_trend(df: pd.DataFrame) -> pd.DataFrame:
    df_temp = df.copy()
    df_temp["hour_dt"] = df_temp["timestamp"].dt.floor("h")
    trend = df_temp.groupby("hour_dt").agg(
        total=("status", "count"),
        failed=("status", lambda x: (x == "Failed").sum())
    ).reset_index()
    trend["failure_rate"] = trend["failed"] / trend["total"] * 100
    return trend

def get_failure_pareto(df: pd.DataFrame) -> pd.DataFrame:
    failed_df = df[df["status"] == "Failed"]
    pareto = failed_df["failure_category"].value_counts().reset_index()
    pareto.columns = ["failure_category", "count"]
    total = pareto["count"].sum()
    pareto["cumulative_pct"] = pareto["count"].cumsum() / total * 100
    return pareto

def detect_outlier_terminals(df: pd.DataFrame, threshold: float = 5.0) -> pd.DataFrame:
    grouped = df.groupby(["terminal_id", "location_name", "city", "vendor"]).agg(
        total=("status", "count"),
        failed=("status", lambda x: (x == "Failed").sum())
    ).reset_index()
    grouped["failure_rate"] = grouped["failed"] / grouped["total"] * 100
    grouped["is_outlier"] = grouped["failure_rate"] > threshold
    return grouped.sort_values(by="failure_rate", ascending=False)

def compute_mtbf(df: pd.DataFrame) -> pd.DataFrame:
    outliers = detect_outlier_terminals(df)
    outlier_ids = outliers[outliers["is_outlier"]]["terminal_id"].tolist()
    
    records = []
    for tid in outlier_ids:
        tdf = df[df["terminal_id"] == tid].sort_values("timestamp")
        n_fails = (tdf["status"] == "Failed").sum()
        if n_fails > 0:
            total_hours = 2160 # 90 days * 24h
            est_downtime = n_fails * 1.5 # approx 1.5h per failure
            mtbf = max((total_hours - est_downtime) / n_fails, 1.0)
            records.append({"terminal_id": tid, "mtbf_hours": round(mtbf, 1), "failures": n_fails})
    return pd.DataFrame(records)

def temporal_heatmap_matrix(df: pd.DataFrame) -> pd.DataFrame:
    df_temp = df.copy()
    df_temp["hour"] = df_temp["timestamp"].dt.hour
    df_temp["day_name"] = df_temp["timestamp"].dt.day_name()
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    pivot = df_temp[df_temp["status"] == "Failed"].pivot_table(
        index="day_name", columns="hour", values="status", aggfunc="count", fill_value=0
    )
    return pivot.reindex(day_order)

def latency_correlation(df: pd.DataFrame):
    valid = df.dropna(subset=["response_time_ms"]).copy()
    valid["is_failed"] = (valid["status"] == "Failed").astype(int)
    r_val, p_val = stats.pearsonr(valid["response_time_ms"], valid["is_failed"])
    return r_val, p_val
