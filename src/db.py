import sqlite3
import pandas as pd
import streamlit as st
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "app.db")

def get_connection():
    abs_path = os.path.abspath(DB_PATH)
    return sqlite3.connect(f"file:{abs_path}?mode=ro", uri=True)

@st.cache_data(ttl=3600)
def load_all_data():
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM v_transaction_analysis", conn)
    conn.close()
    df["timestamp"] = pd.to_datetime(df["timestamp"])
    return df

@st.cache_data(ttl=3600)
def load_terminals():
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM atm_terminals", conn)
    conn.close()
    return df

@st.cache_data(ttl=3600)
def load_maintenance():
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM maintenance_logs", conn)
    conn.close()
    df["event_datetime"] = pd.to_datetime(df["event_datetime"])
    return df
