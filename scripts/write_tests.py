import os

# ─────────────────────────── Step 13 Tests ─────────────────────────────
os.makedirs("tests", exist_ok=True)

conftest_py = '''import pytest
import sqlite3
import pandas as pd
import os

@pytest.fixture(scope="session")
def db_path():
    path = os.path.join(os.path.dirname(__file__), "..", "data", "app.db")
    return os.path.abspath(path)

@pytest.fixture(scope="session")
def conn(db_path):
    assert os.path.exists(db_path), f"DB not found at {db_path}"
    connection = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    yield connection
    connection.close()

@pytest.fixture(scope="session")
def sample_transactions(conn):
    df = pd.read_sql("SELECT * FROM transactions LIMIT 500", conn)
    df["timestamp"] = pd.to_datetime(df["timestamp"])
    return df
'''
with open("tests/conftest.py", "w", encoding="utf-8") as f:
    f.write(conftest_py)

test_data_py = '''import pytest
import sqlite3
import pandas as pd

def test_tc_t_01_db_exists(db_path):
    """TC-T-01: Verify SQLite database exists and is connectable."""
    conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    assert conn is not None
    conn.close()

def test_tc_t_02_tables_and_row_counts(conn):
    """TC-T-02: Verify all 4 tables exist with expected minimum rows."""
    expected = {
        "atm_terminals": 200,
        "error_codes": 30,
        "transactions": 50000,
        "maintenance_logs": 1500
    }
    cursor = conn.cursor()
    for table, min_rows in expected.items():
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        count = cursor.fetchone()[0]
        assert count >= min_rows, f"Table {table} has {count} rows, expected >= {min_rows}"

def test_tc_t_03_primary_keys_unique(conn):
    """TC-T-03: Verify primary keys are unique in masters and fact."""
    cursor = conn.cursor()
    pks = [
        ("atm_terminals", "terminal_id"),
        ("error_codes", "error_code"),
        ("maintenance_logs", "maintenance_id")
    ]
    for table, pk in pks:
        cursor.execute(f"SELECT COUNT({pk}), COUNT(DISTINCT {pk}) FROM {table}")
        total, distinct = cursor.fetchone()
        assert total == distinct, f"Duplicate PKs in {table}.{pk}"

def test_tc_t_04_foreign_key_integrity(conn):
    """TC-T-04: Verify foreign key integrity with zero orphan rows."""
    cursor = conn.cursor()
    # transactions -> atm_terminals
    cursor.execute("""
        SELECT COUNT(*) FROM transactions t 
        LEFT JOIN atm_terminals a ON t.terminal_id = a.terminal_id 
        WHERE a.terminal_id IS NULL
    """)
    assert cursor.fetchone()[0] == 0, "Orphan terminal_id in transactions"

    # transactions -> error_codes
    cursor.execute("""
        SELECT COUNT(*) FROM transactions t 
        LEFT JOIN error_codes e ON t.error_code = e.error_code 
        WHERE t.error_code IS NOT NULL AND e.error_code IS NULL
    """)
    assert cursor.fetchone()[0] == 0, "Orphan error_code in transactions"

def test_tc_t_05_view_exists_and_has_columns(conn):
    """TC-T-05: Verify v_transaction_analysis view exists and has 30+ columns."""
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM v_transaction_analysis LIMIT 1")
    cols = [desc[0] for desc in cursor.description]
    assert len(cols) >= 30, f"View has {len(cols)} columns, expected >= 30"
'''
with open("tests/test_data.py", "w", encoding="utf-8") as f:
    f.write(test_data_py)

test_analysis_py = '''import pytest
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
'''
with open("tests/test_analysis.py", "w", encoding="utf-8") as f:
    f.write(test_analysis_py)

test_app_py = '''import pytest
import os
import sys

def test_tc_t_10_app_files_exist():
    """TC-T-10: Verify app.py and all pages exist."""
    pages = [
        "app.py",
        "src/ui.py",
        "pages/01_Fleet_Overview.py",
        "pages/02_Failure_Category.py",
        "pages/03_Terminal_Health.py",
        "pages/04_Temporal_Patterns.py",
        "pages/05_Network_Latency.py",
        "pages/06_Financial_Impact.py",
    ]
    for p in pages:
        assert os.path.exists(p), f"Missing file: {p}"

def test_tc_t_11_ui_theme_config():
    """TC-T-11: Verify .streamlit/config.toml defines valid primary color."""
    config_path = ".streamlit/config.toml"
    assert os.path.exists(config_path)
    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()
    assert 'primaryColor = "#2B6CB0"' in content
'''
with open("tests/test_app.py", "w", encoding="utf-8") as f:
    f.write(test_app_py)

print("Test suite written in tests/ directory.")
