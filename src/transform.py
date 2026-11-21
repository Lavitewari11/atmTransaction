import pandas as pd

def audit_data_quality(df: pd.DataFrame) -> dict:
    missing_latency = df["response_time_ms"].isna().sum()
    duplicate_txns = df.duplicated(subset=["transaction_id"]).sum()
    outlier_amounts = (df["amount"] > 50000).sum()
    return {
        "missing_latency": int(missing_latency),
        "duplicate_txns": int(duplicate_txns),
        "outlier_amounts": int(outlier_amounts)
    }

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    clean_df = df.drop_duplicates(subset=["transaction_id"]).copy()
    clean_df = clean_df[clean_df["amount"] <= 50000]
    return clean_df
