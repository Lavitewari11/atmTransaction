"""
Step 08 – Database Design Excel Generator
ATM Transaction Failure Analysis (p096) – Lavi Tewari
Generates: docs/08_Database_Design.xlsx  (10 sheets)
"""

import os
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

# ── Output path ────────────────────────────────────────────────────────────────
BASE_DIR = r"c:\Users\ADMIN\OneDrive\Desktop\AtmTransactionAnalysis"
OUT_PATH = os.path.join(BASE_DIR, "docs", "08_Database_Design.xlsx")
os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)

# ── Colour constants ───────────────────────────────────────────────────────────
HDR_FILL   = PatternFill("solid", fgColor="1A365D")
ROW_EVEN   = PatternFill("solid", fgColor="F7FAFC")
ROW_ODD    = PatternFill("solid", fgColor="FFFFFF")
HDR_FONT   = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
DATA_FONT  = Font(name="Calibri", size=9, color="2D3748")
THIN       = Side(style="thin", color="CBD5E0")
BORDER     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.remove(wb.active)   # remove default sheet

# ══════════════════════════════════════════════════════════════════════════════
# Helper utilities
# ══════════════════════════════════════════════════════════════════════════════

def add_sheet(wb, name, tab_color=None):
    ws = wb.create_sheet(title=name)
    if tab_color:
        ws.sheet_properties.tabColor = tab_color
    return ws


def write_header(ws, headers, row=1):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.fill   = HDR_FILL
        c.font   = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_row(ws, data, row, even=True):
    fill = ROW_EVEN if even else ROW_ODD
    for col, v in enumerate(data, 1):
        c = ws.cell(row=row, column=col, value=v)
        c.fill      = fill
        c.font      = DATA_FONT
        c.border    = BORDER
        c.alignment = Alignment(vertical="top", wrap_text=True)


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def auto_filter(ws, headers):
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def freeze(ws, cell="A2"):
    ws.freeze_panes = cell


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 – Overview
# ══════════════════════════════════════════════════════════════════════════════
ws1 = add_sheet(wb, "Overview")

# Title block
ws1.merge_cells("A1:H1")
t = ws1["A1"]
t.value     = "ATM Transaction Failure Analysis – Database Design (Step 08)"
t.font      = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
t.fill      = HDR_FILL
t.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[1].height = 28

desc_rows = [
    ("Project",      "ATM Transaction Failure Analysis (p096)"),
    ("Student",      "Lavi Tewari"),
    ("Step",         "08 – Relational Database Design"),
    ("Technology",   "SQLite 3.45+  |  Python 3.12  |  pandas 2.x  |  openpyxl 3.x"),
    ("DB File",      "data/atm_analysis.db  (SQLite single-file database)"),
    ("Description",  (
        "This document contains the complete relational schema for the ATM "
        "Transaction Failure Analysis project. It covers four tables "
        "(atm_terminals, error_codes, transactions, maintenance_logs) and one "
        "analytical view (v_transaction_analysis). The design supports all 12 "
        "Business Requirements (BR-01 to BR-12) and 8 Operational Questions "
        "(OQ-01 to OQ-08) identified in earlier project steps."
    )),
]
for r, (k, v) in enumerate(desc_rows, 3):
    ws1.cell(row=r, column=1, value=k).font = Font(name="Calibri", size=10, bold=True, color="1A365D")
    ws1.cell(row=r, column=2, value=v).font = Font(name="Calibri", size=10, color="2D3748")
    ws1.merge_cells(f"B{r}:H{r}")

ws1.row_dimensions[8].height = 56  # description row

# Table summary table
hdr = ["Table Name", "Purpose", "Grain", "Planned Rows", "Role"]
write_header(ws1, hdr, row=10)
table_summary = [
    ("atm_terminals",    "Master list of all ATM terminals",                "One row per ATM terminal",             200,    "Master / Dimension"),
    ("error_codes",      "ISO 8583 failure code reference",                 "One row per error code",               30,     "Reference / Dimension"),
    ("transactions",     "Individual ATM transaction events",               "One row per transaction attempt",      50_000, "Fact"),
    ("maintenance_logs", "Terminal maintenance & servicing events",         "One row per maintenance event",        1_500,  "Fact"),
]
for i, row_data in enumerate(table_summary, 11):
    write_row(ws1, row_data, row=i, even=(i % 2 == 0))

set_col_widths(ws1, [22, 38, 38, 15, 20])
freeze(ws1, "A2")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 – Tables_Summary
# ══════════════════════════════════════════════════════════════════════════════
ws2 = add_sheet(wb, "Tables_Summary")
hdr2 = ["Table Name", "Grain", "Primary Key", "Foreign Keys", "Column Count",
        "Planned Rows", "Business Purpose", "Linked OQs"]
write_header(ws2, hdr2)
auto_filter(ws2, hdr2)
freeze(ws2)

rows2 = [
    ("atm_terminals",    "One row per ATM terminal",             "terminal_id",    "—",                                              12, 200,
     "Track physical terminal inventory, location, vendor, and operational status",
     "OQ-01, OQ-04, OQ-05, OQ-07"),
    ("error_codes",      "One row per ISO 8583 error code",      "error_code",     "—",                                              6,  30,
     "Classify failures by category and severity; identify customer-fault vs system-fault errors",
     "OQ-02, OQ-03, OQ-06"),
    ("transactions",     "One row per transaction attempt",      "transaction_id", "terminal_id → atm_terminals\nerror_code → error_codes", 12, 50_000,
     "Core fact table: capture every ATM interaction with result, timing, and failure details",
     "OQ-01 to OQ-08"),
    ("maintenance_logs", "One row per maintenance event",        "maintenance_id", "terminal_id → atm_terminals",                    9, 1_500,
     "Record corrective and preventive maintenance; compute MTBF and maintenance costs",
     "OQ-04, OQ-05, OQ-07"),
]
for i, r in enumerate(rows2, 2):
    write_row(ws2, r, row=i, even=(i % 2 == 0))

set_col_widths(ws2, [22, 32, 18, 36, 14, 14, 44, 24])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 – DD_atm_terminals
# ══════════════════════════════════════════════════════════════════════════════
ws3 = add_sheet(wb, "DD_atm_terminals")
hdr3 = ["Column", "Data Type", "Key", "Nullable", "Allowed Values / Range", "Description", "OQ/BR Supported"]
write_header(ws3, hdr3)
auto_filter(ws3, hdr3)
freeze(ws3)

dd_atm = [
    ("terminal_id",    "TEXT",    "PK",  "NO",  "ATM-XXXX (4-digit)",       "Unique identifier for each ATM terminal",                            "All OQs"),
    ("bank_owner",     "TEXT",    "—",   "NO",  "SBI, HDFC, ICICI, Axis, PNB, BOB, Canara, Union", "Bank that owns/operates the terminal",         "OQ-04, OQ-05, BR-03"),
    ("vendor",         "TEXT",    "—",   "NO",  "NCR, Diebold, Wincor, Hyosung, GRG",              "ATM hardware manufacturer/vendor",             "OQ-05, BR-05"),
    ("terminal_model", "TEXT",    "—",   "YES", "Free text (e.g., NCR 6683, Diebold DN200)",       "Hardware model of the ATM",                    "OQ-05, BR-05"),
    ("install_date",   "DATE",    "—",   "NO",  "YYYY-MM-DD; range 2015-01-01 to 2023-12-31",      "Date the terminal was installed/commissioned", "BR-10"),
    ("location_name",  "TEXT",    "—",   "NO",  "Free text",                                       "Human-readable location name (branch/site)",   "OQ-01, OQ-07"),
    ("city",           "TEXT",    "—",   "NO",  "Free text",                                       "City where terminal is deployed",              "OQ-01, OQ-07"),
    ("state",          "TEXT",    "—",   "NO",  "Indian state names",                              "State where terminal is deployed",             "OQ-07, BR-07"),
    ("zone",           "TEXT",    "—",   "NO",  "Metro / Urban / Semi-Urban",                      "RBI deployment zone classification",           "OQ-07, BR-07"),
    ("status",         "TEXT",    "—",   "NO",  "Active / Inactive",                               "Current operational status of the terminal",   "OQ-04, BR-04"),
    ("latitude",       "REAL",    "—",   "YES", "6.0 – 37.0 (India bounding box)",                "GPS latitude of the terminal",                 "OQ-07"),
    ("longitude",      "REAL",    "—",   "YES", "68.0 – 97.0 (India bounding box)",               "GPS longitude of the terminal",                "OQ-07"),
]
for i, r in enumerate(dd_atm, 2):
    write_row(ws3, r, row=i, even=(i % 2 == 0))
set_col_widths(ws3, [20, 12, 8, 10, 42, 46, 24])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 – DD_error_codes
# ══════════════════════════════════════════════════════════════════════════════
ws4 = add_sheet(wb, "DD_error_codes")
hdr4 = ["Column", "Data Type", "Key", "Nullable", "Allowed Values / Range", "Description", "OQ/BR Supported"]
write_header(ws4, hdr4)
auto_filter(ws4, hdr4)
freeze(ws4)

dd_err = [
    ("error_code",        "TEXT",    "PK", "NO",  "2–4 alphanumeric chars (e.g., D0010, N0001)", "ISO 8583 / proprietary error code",             "OQ-02, OQ-03, BR-02"),
    ("iso_name",          "TEXT",    "—",  "NO",  "Free text (e.g., Decline, Timeout)",          "ISO 8583 standard response name",               "OQ-02, BR-02"),
    ("description",       "TEXT",    "—",  "NO",  "Free text (≤ 255 chars)",                     "Human-readable description of the error",       "OQ-02, BR-02"),
    ("failure_category",  "TEXT",    "—",  "NO",  "Hardware / Network-Switch / Customer-Account / Host-Timeout", "High-level failure category",   "OQ-02, OQ-03, BR-02, BR-11"),
    ("severity",          "TEXT",    "—",  "NO",  "Critical / High / Medium / Low",              "Operational severity of this error type",       "OQ-03, BR-11"),
    ("is_customer_fault", "INTEGER", "—",  "NO",  "0 (system fault) / 1 (customer fault)",       "Flag indicating whether the customer caused the failure", "OQ-06, BR-06"),
]
for i, r in enumerate(dd_err, 2):
    write_row(ws4, r, row=i, even=(i % 2 == 0))
set_col_widths(ws4, [20, 12, 8, 10, 48, 44, 24])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 – DD_transactions
# ══════════════════════════════════════════════════════════════════════════════
ws5 = add_sheet(wb, "DD_transactions")
hdr5 = ["Column", "Data Type", "Key", "Nullable", "Allowed Values / Range", "Description", "OQ/BR Supported"]
write_header(ws5, hdr5)
auto_filter(ws5, hdr5)
freeze(ws5)

dd_txn = [
    ("transaction_id",   "TEXT",     "PK",           "NO",  "TXN-XXXXXXXX (8-char hex suffix)",              "Unique transaction identifier",                     "All OQs"),
    ("terminal_id",      "TEXT",     "FK→atm_terminals", "NO", "Must exist in atm_terminals.terminal_id",   "ATM terminal where transaction occurred",           "OQ-01, OQ-04, OQ-05, BR-01"),
    ("timestamp",        "DATETIME", "—",            "NO",  "YYYY-MM-DD HH:MM:SS; 2023-01-01 to 2024-12-31","Date and time of the transaction attempt",          "OQ-01, OQ-08, BR-01, BR-08"),
    ("transaction_type", "TEXT",     "—",            "NO",  "Cash-Withdrawal / Balance-Inquiry / PIN-Change / Mini-Statement", "Type of ATM transaction attempted", "OQ-06, BR-06"),
    ("amount",           "REAL",     "—",            "YES", "0 – 50000 (INR); NULL for non-cash txns",       "Transaction amount in Indian Rupees",               "OQ-08, BR-08"),
    ("status",           "TEXT",     "—",            "NO",  "Approved / Failed / Reversed",                  "Final outcome of the transaction",                  "All OQs"),
    ("error_code",       "TEXT",     "FK→error_codes", "YES","Must exist in error_codes.error_code or NULL", "Error code if transaction failed; NULL if approved", "OQ-02, OQ-03, BR-02"),
    ("response_time_ms", "INTEGER",  "—",            "NO",  "100 – 30000 (milliseconds)",                    "Host response latency in milliseconds",             "OQ-03, BR-03"),
    ("card_type",        "TEXT",     "—",            "NO",  "On-Us / Off-Us",                                "Whether card issued by ATM-owning bank (On-Us) or another bank (Off-Us)", "OQ-06, BR-06"),
    ("issuing_bank",     "TEXT",     "—",            "NO",  "Bank name (SBI, HDFC, ICICI, Axis, etc.)",      "Bank that issued the transaction card",             "OQ-06, BR-06"),
    ("failure_category", "TEXT",     "—",            "YES", "Hardware / Network-Switch / Customer-Account / Host-Timeout / NULL", "Denormalised failure category for fast querying", "OQ-02, OQ-03, BR-11"),
    ("failure_reason",   "TEXT",     "—",            "YES", "Free text ≤ 200 chars / NULL",                  "Short description of why the transaction failed",   "OQ-02, BR-02"),
]
for i, r in enumerate(dd_txn, 2):
    write_row(ws5, r, row=i, even=(i % 2 == 0))
set_col_widths(ws5, [20, 12, 20, 10, 50, 44, 24])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6 – DD_maintenance_logs
# ══════════════════════════════════════════════════════════════════════════════
ws6 = add_sheet(wb, "DD_maintenance_logs")
hdr6 = ["Column", "Data Type", "Key", "Nullable", "Allowed Values / Range", "Description", "OQ/BR Supported"]
write_header(ws6, hdr6)
auto_filter(ws6, hdr6)
freeze(ws6)

dd_maint = [
    ("maintenance_id",     "TEXT",     "PK",               "NO",  "MNT-XXXXXXXX",                                     "Unique maintenance event identifier",             "OQ-04, OQ-05"),
    ("terminal_id",        "TEXT",     "FK→atm_terminals", "NO",  "Must exist in atm_terminals.terminal_id",          "ATM terminal serviced",                           "OQ-04, OQ-05, BR-04"),
    ("event_datetime",     "DATETIME", "—",                "NO",  "YYYY-MM-DD HH:MM:SS",                              "Date and time the maintenance event started",     "OQ-04, OQ-05, BR-10"),
    ("event_type",         "TEXT",     "—",                "NO",  "Preventive / Corrective / Emergency",              "Category of maintenance activity",                "OQ-04, BR-04"),
    ("component_replaced", "TEXT",     "—",                "YES", "Free text / NULL",                                 "Name of hardware component replaced, if any",     "OQ-05, BR-05"),
    ("engineer_id",        "TEXT",     "—",                "NO",  "ENG-XXXX",                                         "Identifier of field engineer who performed work",  "OQ-04"),
    ("resolution_notes",   "TEXT",     "—",                "YES", "Free text ≤ 500 chars",                            "Description of work performed and resolution",    "OQ-04"),
    ("duration_hours",     "REAL",     "—",                "NO",  "0.25 – 48.0 (hours)",                             "Total hours the terminal was out of service",      "OQ-04, BR-04, BR-10"),
    ("cost_inr",           "REAL",     "—",                "NO",  "500 – 500000 (INR)",                               "Total cost of the maintenance event in INR",      "OQ-05, BR-05"),
]
for i, r in enumerate(dd_maint, 2):
    write_row(ws6, r, row=i, even=(i % 2 == 0))
set_col_widths(ws6, [22, 12, 20, 10, 42, 44, 24])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7 – Relationships
# ══════════════════════════════════════════════════════════════════════════════
ws7 = add_sheet(wb, "Relationships")
hdr7 = ["Parent Table", "Child Table", "Join Key", "Cardinality", "Relationship Rule", "Business Meaning"]
write_header(ws7, hdr7)
auto_filter(ws7, hdr7)
freeze(ws7)

rels = [
    ("atm_terminals", "transactions",     "terminal_id",  "1 : Many",
     "Every transaction MUST reference a valid terminal_id in atm_terminals",
     "Links each transaction to the physical ATM terminal where it occurred"),
    ("atm_terminals", "maintenance_logs", "terminal_id",  "1 : Many",
     "Every maintenance event MUST reference a valid terminal_id in atm_terminals",
     "Links each maintenance event to the serviced terminal for MTBF analysis"),
    ("error_codes",   "transactions",     "error_code",   "1 : Many (optional)",
     "A failed/reversed transaction MAY reference a valid error_code; approved transactions set error_code = NULL",
     "Classifies transaction failures by ISO 8583 error category and severity"),
]
for i, r in enumerate(rels, 2):
    write_row(ws7, r, row=i, even=(i % 2 == 0))
set_col_widths(ws7, [22, 22, 18, 16, 48, 48])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8 – KPIs_Measures
# ══════════════════════════════════════════════════════════════════════════════
ws8 = add_sheet(wb, "KPIs_Measures")
hdr8 = ["Measure Name", "Formula (SQL or pseudo)", "OQ Answered", "BR Supported", "Output Type"]
write_header(ws8, hdr8)
auto_filter(ws8, hdr8)
freeze(ws8)

kpis = [
    ("Fleet Failure Rate (%)",
     "COUNT(CASE WHEN status='Failed' THEN 1 END) * 100.0 / COUNT(*)",
     "OQ-01", "BR-01", "Percentage (0–100)"),

    ("Terminal Failure Rate – 7-day rolling",
     ("SELECT terminal_id,\n"
      "  SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END) OVER\n"
      "    (PARTITION BY terminal_id ORDER BY DATE(timestamp)\n"
      "     RANGE BETWEEN 6 PRECEDING AND CURRENT ROW) * 100.0 /\n"
      "  COUNT(*) OVER\n"
      "    (PARTITION BY terminal_id ORDER BY DATE(timestamp)\n"
      "     RANGE BETWEEN 6 PRECEDING AND CURRENT ROW) AS rolling_failure_rate\n"
      "FROM transactions"),
     "OQ-01, OQ-04", "BR-01, BR-04", "Percentage (rolling window)"),

    ("MTBF (hours)",
     ("(SUM(window_hours) - SUM(downtime_hours)) / NULLIF(COUNT(failure_events), 0)\n"
      "-- window_hours from date range; downtime_hours from maintenance_logs.duration_hours"),
     "OQ-04", "BR-04, BR-10", "Hours (REAL)"),

    ("Latency p50 (ms)",
     "PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY response_time_ms)",
     "OQ-03", "BR-03", "Milliseconds (INTEGER)"),

    ("Latency p90 (ms)",
     "PERCENTILE_CONT(0.9) WITHIN GROUP (ORDER BY response_time_ms)",
     "OQ-03", "BR-03", "Milliseconds (INTEGER)"),

    ("Latency p99 (ms)",
     "PERCENTILE_CONT(0.99) WITHIN GROUP (ORDER BY response_time_ms)",
     "OQ-03", "BR-03", "Milliseconds (INTEGER)"),

    ("Financial Disruption (INR)",
     "SUM(amount) FILTER (WHERE status = 'Failed')",
     "OQ-08", "BR-08", "Currency – INR (REAL)"),

    ("Vendor Failure Rate per 1000",
     ("COUNT(CASE WHEN t.status='Failed' THEN 1 END) * 1000.0 / COUNT(*)\n"
      "-- grouped by atm_terminals.vendor"),
     "OQ-05", "BR-05", "Rate per 1000 transactions"),

    ("On-Us vs Off-Us Failure Rate Differential",
     ("COUNT(CASE WHEN card_type='On-Us'  AND status='Failed' THEN 1 END)*100.0/NULLIF(COUNT(CASE WHEN card_type='On-Us'  THEN 1 END),0)\n"
      "MINUS\n"
      "COUNT(CASE WHEN card_type='Off-Us' AND status='Failed' THEN 1 END)*100.0/NULLIF(COUNT(CASE WHEN card_type='Off-Us' THEN 1 END),0)"),
     "OQ-06", "BR-06", "Percentage differential"),

    ("Hardware Failure Share (%)",
     "COUNT(CASE WHEN failure_category='Hardware' THEN 1 END)*100.0 / NULLIF(COUNT(CASE WHEN status='Failed' THEN 1 END),0)",
     "OQ-02", "BR-02, BR-11", "Percentage"),

    ("Network/Switch Failure Share (%)",
     "COUNT(CASE WHEN failure_category='Network-Switch' THEN 1 END)*100.0 / NULLIF(COUNT(CASE WHEN status='Failed' THEN 1 END),0)",
     "OQ-02", "BR-02, BR-11", "Percentage"),

    ("Customer-Fault Failure Share (%)",
     "COUNT(CASE WHEN ec.is_customer_fault=1 THEN 1 END)*100.0 / NULLIF(COUNT(CASE WHEN t.status='Failed' THEN 1 END),0)",
     "OQ-06", "BR-06", "Percentage"),

    ("Zone-wise Failure Rate (%)",
     ("COUNT(CASE WHEN t.status='Failed' THEN 1 END)*100.0/COUNT(*)\n"
      "-- GROUP BY a.zone"),
     "OQ-07", "BR-07", "Percentage by zone"),

    ("Monthly Failure Trend",
     ("SELECT strftime('%Y-%m', timestamp) AS month,\n"
      "  COUNT(CASE WHEN status='Failed' THEN 1 END)*100.0/COUNT(*) AS failure_rate\n"
      "FROM transactions GROUP BY 1 ORDER BY 1"),
     "OQ-01, OQ-08", "BR-01, BR-08", "Time series percentage"),
]
for i, r in enumerate(kpis, 2):
    write_row(ws8, r, row=i, even=(i % 2 == 0))
set_col_widths(ws8, [32, 72, 20, 18, 26])


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 9 – SQLite_DDL   (tab colour = red)
# ══════════════════════════════════════════════════════════════════════════════
ws9 = add_sheet(wb, "SQLite_DDL", tab_color="E53E3E")

DDL = """\
-- ============================================================
--  ATM Transaction Failure Analysis – SQLite DDL
--  Database: data/atm_analysis.db
--  SQLite version: 3.45+
--  Author: Lavi Tewari  |  Step 08
-- ============================================================

PRAGMA foreign_keys = ON;

-- ─────────────────────────────────────────────────────────────
-- TABLE 1 : atm_terminals  (Master / Dimension)
-- ─────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS atm_terminals (
    terminal_id    TEXT NOT NULL,
    bank_owner     TEXT NOT NULL,
    vendor         TEXT NOT NULL
        CHECK (vendor IN ('NCR','Diebold','Wincor','Hyosung','GRG')),
    terminal_model TEXT,
    install_date   DATE NOT NULL,
    location_name  TEXT NOT NULL,
    city           TEXT NOT NULL,
    state          TEXT NOT NULL,
    zone           TEXT NOT NULL
        CHECK (zone IN ('Metro','Urban','Semi-Urban')),
    status         TEXT NOT NULL DEFAULT 'Active'
        CHECK (status IN ('Active','Inactive')),
    latitude       REAL,
    longitude      REAL,
    CONSTRAINT pk_atm_terminals PRIMARY KEY (terminal_id)
);

CREATE INDEX IF NOT EXISTS idx_atm_terminals_zone
    ON atm_terminals (zone);
CREATE INDEX IF NOT EXISTS idx_atm_terminals_status
    ON atm_terminals (status);
CREATE INDEX IF NOT EXISTS idx_atm_terminals_vendor
    ON atm_terminals (vendor);
CREATE INDEX IF NOT EXISTS idx_atm_terminals_bank
    ON atm_terminals (bank_owner);

-- ─────────────────────────────────────────────────────────────
-- TABLE 2 : error_codes  (Reference / Dimension)
-- ─────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS error_codes (
    error_code        TEXT NOT NULL,
    iso_name          TEXT NOT NULL,
    description       TEXT NOT NULL,
    failure_category  TEXT NOT NULL
        CHECK (failure_category IN (
            'Hardware','Network-Switch','Customer-Account','Host-Timeout')),
    severity          TEXT NOT NULL
        CHECK (severity IN ('Critical','High','Medium','Low')),
    is_customer_fault INTEGER NOT NULL DEFAULT 0
        CHECK (is_customer_fault IN (0, 1)),
    CONSTRAINT pk_error_codes PRIMARY KEY (error_code)
);

CREATE INDEX IF NOT EXISTS idx_error_codes_category
    ON error_codes (failure_category);
CREATE INDEX IF NOT EXISTS idx_error_codes_severity
    ON error_codes (severity);

-- ─────────────────────────────────────────────────────────────
-- TABLE 3 : transactions  (Fact – 50 000 rows)
-- ─────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS transactions (
    transaction_id   TEXT    NOT NULL,
    terminal_id      TEXT    NOT NULL,
    timestamp        DATETIME NOT NULL,
    transaction_type TEXT    NOT NULL
        CHECK (transaction_type IN (
            'Cash-Withdrawal','Balance-Inquiry','PIN-Change','Mini-Statement')),
    amount           REAL,
    status           TEXT    NOT NULL
        CHECK (status IN ('Approved','Failed','Reversed')),
    error_code       TEXT,
    response_time_ms INTEGER NOT NULL
        CHECK (response_time_ms > 0),
    card_type        TEXT    NOT NULL
        CHECK (card_type IN ('On-Us','Off-Us')),
    issuing_bank     TEXT    NOT NULL,
    failure_category TEXT,
    failure_reason   TEXT,
    CONSTRAINT pk_transactions    PRIMARY KEY (transaction_id),
    CONSTRAINT fk_txn_terminal    FOREIGN KEY (terminal_id)
        REFERENCES atm_terminals (terminal_id)
        ON UPDATE CASCADE ON DELETE RESTRICT,
    CONSTRAINT fk_txn_error_code  FOREIGN KEY (error_code)
        REFERENCES error_codes (error_code)
        ON UPDATE CASCADE ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_txn_terminal_id
    ON transactions (terminal_id);
CREATE INDEX IF NOT EXISTS idx_txn_timestamp
    ON transactions (timestamp);
CREATE INDEX IF NOT EXISTS idx_txn_status
    ON transactions (status);
CREATE INDEX IF NOT EXISTS idx_txn_error_code
    ON transactions (error_code);
CREATE INDEX IF NOT EXISTS idx_txn_card_type
    ON transactions (card_type);
CREATE INDEX IF NOT EXISTS idx_txn_failure_category
    ON transactions (failure_category);
-- Composite index for time-series failure-rate queries
CREATE INDEX IF NOT EXISTS idx_txn_terminal_ts
    ON transactions (terminal_id, timestamp);

-- ─────────────────────────────────────────────────────────────
-- TABLE 4 : maintenance_logs  (Fact – 1 500 rows)
-- ─────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS maintenance_logs (
    maintenance_id     TEXT     NOT NULL,
    terminal_id        TEXT     NOT NULL,
    event_datetime     DATETIME NOT NULL,
    event_type         TEXT     NOT NULL
        CHECK (event_type IN ('Preventive','Corrective','Emergency')),
    component_replaced TEXT,
    engineer_id        TEXT     NOT NULL,
    resolution_notes   TEXT,
    duration_hours     REAL     NOT NULL
        CHECK (duration_hours > 0),
    cost_inr           REAL     NOT NULL
        CHECK (cost_inr >= 0),
    CONSTRAINT pk_maintenance      PRIMARY KEY (maintenance_id),
    CONSTRAINT fk_maint_terminal   FOREIGN KEY (terminal_id)
        REFERENCES atm_terminals (terminal_id)
        ON UPDATE CASCADE ON DELETE RESTRICT
);

CREATE INDEX IF NOT EXISTS idx_maint_terminal_id
    ON maintenance_logs (terminal_id);
CREATE INDEX IF NOT EXISTS idx_maint_event_datetime
    ON maintenance_logs (event_datetime);
CREATE INDEX IF NOT EXISTS idx_maint_event_type
    ON maintenance_logs (event_type);

-- ─────────────────────────────────────────────────────────────
-- VIEW : v_transaction_analysis
--   Flattened join of all 4 tables for BI / reporting queries
-- ─────────────────────────────────────────────────────────────
CREATE VIEW IF NOT EXISTS v_transaction_analysis AS
SELECT
    -- Transaction fields
    t.transaction_id,
    t.timestamp,
    strftime('%Y',  t.timestamp)         AS txn_year,
    strftime('%m',  t.timestamp)         AS txn_month,
    strftime('%Y-%m', t.timestamp)       AS txn_year_month,
    strftime('%w',  t.timestamp)         AS txn_day_of_week,
    strftime('%H',  t.timestamp)         AS txn_hour,
    t.transaction_type,
    t.amount,
    t.status,
    t.response_time_ms,
    t.card_type,
    t.issuing_bank,
    t.failure_category   AS txn_failure_category,
    t.failure_reason,

    -- Terminal fields
    a.terminal_id,
    a.bank_owner,
    a.vendor,
    a.terminal_model,
    a.install_date,
    a.location_name,
    a.city,
    a.state,
    a.zone,
    a.status             AS terminal_status,
    a.latitude,
    a.longitude,

    -- Error-code fields (NULL when transaction approved)
    ec.error_code,
    ec.iso_name,
    ec.description       AS error_description,
    ec.failure_category  AS ec_failure_category,
    ec.severity,
    ec.is_customer_fault,

    -- Derived / computed columns
    CASE WHEN t.status = 'Failed' THEN 1 ELSE 0 END       AS is_failed,
    CASE WHEN t.status = 'Approved' THEN t.amount ELSE 0 END AS approved_amount,
    CASE WHEN t.status = 'Failed'   THEN t.amount ELSE 0 END AS disrupted_amount,

    -- Maintenance summary per terminal (sub-select for lightweight join)
    (SELECT COUNT(*)
     FROM maintenance_logs ml
     WHERE ml.terminal_id = a.terminal_id)                 AS total_maintenance_events,
    (SELECT COALESCE(SUM(ml.duration_hours), 0)
     FROM maintenance_logs ml
     WHERE ml.terminal_id = a.terminal_id)                 AS total_downtime_hours,
    (SELECT COALESCE(SUM(ml.cost_inr), 0)
     FROM maintenance_logs ml
     WHERE ml.terminal_id = a.terminal_id)                 AS total_maintenance_cost_inr

FROM transactions t
JOIN  atm_terminals a  ON a.terminal_id = t.terminal_id
LEFT JOIN error_codes ec ON ec.error_code = t.error_code;

-- ─────────────────────────────────────────────────────────────
-- END OF DDL
-- ─────────────────────────────────────────────────────────────
"""

ws9.column_dimensions["A"].width = 110
ws9["A1"] = DDL
ws9["A1"].font      = Font(name="Courier New", size=9, color="1A365D")
ws9["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws9.row_dimensions[1].height = 1200


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 10 – Updated_RTM   (tab colour = green)
# ══════════════════════════════════════════════════════════════════════════════
ws10 = add_sheet(wb, "Updated_RTM", tab_color="38A169")
hdr10 = [
    "BR ID", "Business Requirement", "Priority", "Source",
    "Linked OQ(s)", "KPI / Measure", "Data Source",
    "DB Table.Column", "Status"
]
write_header(ws10, hdr10)
auto_filter(ws10, hdr10)
freeze(ws10)

rtm_rows = [
    ("BR-01", "Track overall ATM fleet failure rate across time",
     "High", "Risk Mgmt", "OQ-01",
     "Fleet Failure Rate (%); Monthly Failure Trend",
     "transactions",
     "transactions.status, transactions.timestamp",
     "Mapped"),

    ("BR-02", "Classify failures by ISO 8583 error code and category",
     "High", "Operations", "OQ-02",
     "Hardware/Network/Customer Failure Share (%)",
     "transactions, error_codes",
     "error_codes.failure_category, error_codes.iso_name, transactions.error_code",
     "Mapped"),

    ("BR-03", "Monitor host response latency (p50/p90/p99)",
     "High", "IT / Network", "OQ-03",
     "Latency p50, p90, p99 (ms)",
     "transactions",
     "transactions.response_time_ms",
     "Mapped"),

    ("BR-04", "Identify high-failure terminals for proactive maintenance",
     "High", "Field Ops", "OQ-04",
     "Terminal Failure Rate (7-day rolling); MTBF",
     "transactions, maintenance_logs, atm_terminals",
     "transactions.terminal_id, maintenance_logs.duration_hours, maintenance_logs.event_type",
     "Mapped"),

    ("BR-05", "Benchmark failure rates and maintenance costs by ATM vendor",
     "Medium", "Procurement", "OQ-05",
     "Vendor Failure Rate per 1000; Vendor Maintenance Cost",
     "atm_terminals, transactions, maintenance_logs",
     "atm_terminals.vendor, maintenance_logs.cost_inr",
     "Mapped"),

    ("BR-06", "Compare On-Us vs Off-Us card failure patterns",
     "Medium", "Card Ops", "OQ-06",
     "On-Us vs Off-Us Failure Rate Differential; Customer-Fault Share",
     "transactions, error_codes",
     "transactions.card_type, transactions.issuing_bank, error_codes.is_customer_fault",
     "Mapped"),

    ("BR-07", "Analyse geographic (zone/state) failure distribution",
     "Medium", "Network Planning", "OQ-07",
     "Zone-wise Failure Rate (%)",
     "atm_terminals, transactions",
     "atm_terminals.zone, atm_terminals.state, atm_terminals.city",
     "Mapped"),

    ("BR-08", "Quantify financial disruption from failed transactions",
     "High", "Finance", "OQ-08",
     "Financial Disruption (INR)",
     "transactions",
     "transactions.amount, transactions.status",
     "Mapped"),

    ("BR-09", "Maintain an error-code master reference for all failure types",
     "Low", "Data Governance", "OQ-02",
     "Error code lookup / description",
     "error_codes",
     "error_codes.error_code, error_codes.description, error_codes.severity",
     "Mapped"),

    ("BR-10", "Calculate MTBF and downtime per terminal over rolling windows",
     "High", "SLA Mgmt", "OQ-04",
     "MTBF (hours); Total Downtime Hours",
     "maintenance_logs, atm_terminals",
     "maintenance_logs.duration_hours, maintenance_logs.event_datetime, atm_terminals.install_date",
     "Mapped"),

    ("BR-11", "Rank failure severity (Critical/High/Medium/Low) for triage",
     "High", "Operations", "OQ-02, OQ-03",
     "Failure severity distribution by category",
     "error_codes, transactions",
     "error_codes.severity, error_codes.failure_category, transactions.failure_category",
     "Mapped"),

    ("BR-12", "Support drill-down from fleet level to individual terminal",
     "Medium", "Executive Reporting", "OQ-01 to OQ-08",
     "All KPIs aggregated at terminal / zone / vendor level",
     "v_transaction_analysis",
     "v_transaction_analysis.* (all 35 columns)",
     "Mapped"),
]

for i, r in enumerate(rtm_rows, 2):
    write_row(ws10, r, row=i, even=(i % 2 == 0))

set_col_widths(ws10, [8, 44, 10, 18, 18, 42, 32, 48, 10])


# ══════════════════════════════════════════════════════════════════════════════
# Save workbook
# ══════════════════════════════════════════════════════════════════════════════
wb.save(OUT_PATH)
print(f"SUCCESS – workbook saved to: {OUT_PATH}")
print(f"Sheets created ({len(wb.sheetnames)}):", wb.sheetnames)
