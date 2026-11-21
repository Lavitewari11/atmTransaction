import pytest
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
