import pytest
import pandas as pd
import numpy as np

def test_tc_t_06_failure_rate_formula():
    """TC-T-06: Verify fleet failure rate mathematical calculation."""
    df = pd.DataFrame({
        "status": ["Approved", "Failed", "Approved", "Failed", "Approved"]
    })
    rate = (df["status"] == "Failed").sum() / len(df) * 100
    assert rate == 40.0

def test_tc_t_07_outlier_detection_logic():
    """TC-T-07: Verify outlier detection flags terminal > 5% failure threshold."""
    terminals_df = pd.DataFrame({
        "terminal_id": ["ATM_001", "ATM_002"],
        "total": [100, 100],
        "failed": [6, 2]  # 6% vs 2%
    })
    terminals_df["failure_rate"] = terminals_df["failed"] / terminals_df["total"] * 100
    outliers = terminals_df[terminals_df["failure_rate"] > 5.0]["terminal_id"].tolist()
    assert outliers == ["ATM_001"]

def test_tc_t_08_mtbf_formula():
    """TC-T-08: Verify MTBF = (window_hours - downtime_hours) / failures."""
    window_hours = 2160  # 90 days * 24h
    downtime_hours = 160
    failures = 50
    mtbf = (window_hours - downtime_hours) / failures
    assert mtbf == 40.0

def test_tc_t_09_financial_disruption_sum():
    """TC-T-09: Verify financial disruption aggregates only Failed records."""
    df = pd.DataFrame({
        "amount": [5000, 2000, 10000, 3000],
        "status": ["Approved", "Failed", "Failed", "Approved"]
    })
    disrupted = df[df["status"] == "Failed"]["amount"].sum()
    assert disrupted == 12000
