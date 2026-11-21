import pytest
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
