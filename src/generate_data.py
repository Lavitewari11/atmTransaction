"""
ATM Transaction Failure Analysis — Data Generation Script
Project: p096 | Student: Lavi Tewari
Step 10: Synthetic data generation with embedded patterns & data quality issues

Run from project root:
    python src/generate_data.py
"""

import os
import sqlite3
import warnings
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# 0. SEED & OUTPUT PATHS
# ─────────────────────────────────────────────
np.random.seed(42)
rng = np.random.default_rng(42)

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
os.makedirs(DATA_DIR, exist_ok=True)

EXCEL_PATH  = os.path.join(DATA_DIR, "atm_transactions_data.xlsx")
CSV_TERMS   = os.path.join(DATA_DIR, "atm_terminals.csv")
CSV_ERRS    = os.path.join(DATA_DIR, "error_codes.csv")
CSV_TXN     = os.path.join(DATA_DIR, "transactions.csv")
CSV_MAINT   = os.path.join(DATA_DIR, "maintenance_logs.csv")
DB_PATH     = os.path.join(DATA_DIR, "app.db")

print("=" * 65)
print("  ATM Transaction Failure Analysis — Data Generator")
print("=" * 65)

# ═══════════════════════════════════════════════════════════════
# TABLE 1: atm_terminals  (200 rows)
# ═══════════════════════════════════════════════════════════════
print("\n[1/4] Generating atm_terminals …")

BANKS   = ["SBI", "HDFC", "ICICI", "PNB", "BOI"]
VENDORS = ["NCR", "Diebold", "Hitachi", "GRGBanking", "Nautilus"]
VENDOR_MODELS = {
    "NCR":        "NCR-6683",
    "Diebold":    "Diebold-DN200",
    "Hitachi":    "Hitachi-2845",
    "GRGBanking": "GRG-H68N",
    "Nautilus":   "Nautilus-MX5600",
}

# 20 Indian cities with state, zone, lat, lon
CITY_DATA = [
    # Metro cities
    ("Mumbai",      "Maharashtra",      "Metro",      19.0760,  72.8777),
    ("Delhi",       "Delhi",            "Metro",      28.7041,  77.1025),
    ("Bengaluru",   "Karnataka",        "Metro",      12.9716,  77.5946),
    ("Chennai",     "Tamil Nadu",       "Metro",      13.0827,  80.2707),
    ("Kolkata",     "West Bengal",      "Metro",      22.5726,  88.3639),
    ("Hyderabad",   "Telangana",        "Metro",      17.3850,  78.4867),
    # Urban cities
    ("Pune",        "Maharashtra",      "Urban",      18.5204,  73.8567),
    ("Ahmedabad",   "Gujarat",          "Urban",      23.0225,  72.5714),
    ("Jaipur",      "Rajasthan",        "Urban",      26.9124,  75.7873),
    ("Lucknow",     "Uttar Pradesh",    "Urban",      26.8467,  80.9462),
    ("Surat",       "Gujarat",          "Urban",      21.1702,  72.8311),
    ("Nagpur",      "Maharashtra",      "Urban",      21.1458,  79.0882),
    # Semi-Urban cities
    ("Bhopal",      "Madhya Pradesh",   "Semi-Urban", 23.2599,  77.4126),
    ("Patna",       "Bihar",            "Semi-Urban", 25.5941,  85.1376),
    ("Chandigarh",  "Punjab",           "Semi-Urban", 30.7333,  76.7794),
    ("Coimbatore",  "Tamil Nadu",       "Semi-Urban", 11.0168,  76.9558),
    ("Vadodara",    "Gujarat",          "Semi-Urban", 22.3072,  73.1812),
    ("Indore",      "Madhya Pradesh",   "Semi-Urban", 22.7196,  75.8577),
    ("Kochi",       "Kerala",           "Semi-Urban", 10.0761,  76.2920),
    ("Visakhapatnam","Andhra Pradesh",  "Semi-Urban", 17.6868,  83.2185),
]

LOCATION_SUFFIXES = [
    "Branch", "Kiosk", "ATM Point", "Service Centre", "Mall Outlet",
    "Railway Station", "Airport Terminal", "Market Branch", "Campus ATM",
    "Hospital Wing",
]
LOCATION_PREFIXES = [
    "Connaught Place", "Bandra West", "Koramangala", "Anna Nagar",
    "Salt Lake", "Jubilee Hills", "Viman Nagar", "Navrangpura",
    "Vaishali Nagar", "Gomti Nagar", "Adajan", "Dharampeth",
    "Arera Colony", "Boring Road", "Sector 17", "RS Puram",
    "Alkapuri", "Vijay Nagar", "Palarivattom", "Dwaraka Nagar",
    "MG Road", "Brigade Road", "Park Street", "Linking Road",
    "Civil Lines", "Ashok Nagar", "Tilak Nagar", "Lajpat Nagar",
    "Karol Bagh", "Saket", "Powai", "Andheri East",
]

n_terminals = 200
terminal_ids = [f"ATM_{str(i).zfill(3)}" for i in range(1, n_terminals + 1)]

# Distribute banks ~40 each
bank_col   = np.tile(BANKS, 40)                              # 200 exact
np.random.shuffle(bank_col)

# Distribute vendors ~40 each
vendor_col = np.tile(VENDORS, 40)
np.random.shuffle(vendor_col)

# Install dates between 2018-01-01 and 2023-12-31
install_start = datetime(2018, 1, 1).toordinal()
install_end   = datetime(2023, 12, 31).toordinal()
install_dates = [
    datetime.fromordinal(np.random.randint(install_start, install_end + 1)).strftime("%Y-%m-%d")
    for _ in range(n_terminals)
]

# City assignment
city_indices = np.random.randint(0, len(CITY_DATA), n_terminals)
cities      = [CITY_DATA[i][0] for i in city_indices]
states      = [CITY_DATA[i][1] for i in city_indices]
zones       = [CITY_DATA[i][2] for i in city_indices]
base_lats   = [CITY_DATA[i][3] for i in city_indices]
base_lons   = [CITY_DATA[i][4] for i in city_indices]

# Add small jitter to lat/lon
lat_col = np.round([base_lats[i] + np.random.uniform(-0.05, 0.05) for i in range(n_terminals)], 6)
lon_col = np.round([base_lons[i] + np.random.uniform(-0.05, 0.05) for i in range(n_terminals)], 6)

# Location names
loc_prefixes = np.random.choice(LOCATION_PREFIXES, n_terminals, replace=True)
loc_suffixes = np.random.choice(LOCATION_SUFFIXES, n_terminals, replace=True)
location_names = [f"{p} {s}" for p, s in zip(loc_prefixes, loc_suffixes)]

# Status: 90% Active, 10% Inactive
status_col = np.where(np.random.random(n_terminals) < 0.90, "Active", "Inactive")

terminals_df = pd.DataFrame({
    "terminal_id":    terminal_ids,
    "bank_owner":     bank_col,
    "vendor":         vendor_col,
    "terminal_model": [VENDOR_MODELS[v] for v in vendor_col],
    "install_date":   install_dates,
    "location_name":  location_names,
    "city":           cities,
    "state":          states,
    "zone":           zones,
    "status":         status_col,
    "latitude":       lat_col,
    "longitude":      lon_col,
})
print(f"    → {len(terminals_df)} rows generated.")

# ═══════════════════════════════════════════════════════════════
# TABLE 2: error_codes  (30 rows)
# ═══════════════════════════════════════════════════════════════
print("\n[2/4] Generating error_codes …")

# Format: code, description, category, severity
ERROR_CODE_DATA = [
    ("00", "Approved",                        "None",             "None",     0),
    ("51", "Insufficient Funds",              "Customer-Account", "Medium",   1),
    ("55", "Incorrect PIN",                   "Customer-Account", "Low",      1),
    ("57", "Transaction Not Permitted",       "Customer-Account", "Medium",   1),
    ("61", "Withdrawal Limit Exceeded",       "Customer-Account", "Low",      1),
    ("65", "Withdrawal Frequency Limit",      "Customer-Account", "Low",      1),
    ("91", "Issuer Inoperative",              "Host-Timeout",     "Critical", 0),
    ("96", "System Malfunction",              "Host-Timeout",     "Critical", 0),
    ("68", "Response Received Too Late",      "Network-Switch",   "High",     0),
    ("05", "Do Not Honour",                   "Host-Timeout",     "High",     0),
    ("14", "Invalid Card Number",             "Customer-Account", "Medium",   1),
    ("41", "Lost Card",                       "Customer-Account", "High",     1),
    ("43", "Stolen Card",                     "Customer-Account", "High",     1),
    ("54", "Expired Card",                    "Customer-Account", "Low",      1),
    ("62", "Restricted Card",                 "Customer-Account", "Medium",   1),
    ("76", "Invalid Account",                 "Customer-Account", "Medium",   1),
    ("85", "No Reason to Decline (Reversal)", "Customer-Account", "Low",      1),
    ("34", "Fraud Suspicion",                 "Customer-Account", "Critical", 1),
    ("13", "Invalid Amount",                  "Customer-Account", "Low",      1),
    ("30", "Format Error",                    "Network-Switch",   "Medium",   0),
    ("31", "Bank Not Supported",              "Network-Switch",   "Medium",   0),
    ("44", "No Account",                      "Customer-Account", "Medium",   1),
    ("52", "No Checking Account",             "Customer-Account", "Low",      1),
    ("53", "No Savings Account",              "Customer-Account", "Low",      1),
    ("58", "Terminal Not Permitted",          "Hardware",         "High",     0),
    ("75", "PIN Tries Exceeded",              "Customer-Account", "Medium",   1),
    ("P1", "Cash Dispenser Jam",              "Hardware",         "Critical", 0),
    ("P2", "Card Reader Failure",             "Hardware",         "High",     0),
    ("P3", "Receipt Printer Error",           "Hardware",         "Low",      0),
    ("N1", "Network Link Down",               "Network-Switch",   "Critical", 0),
]

error_codes_df = pd.DataFrame(ERROR_CODE_DATA, columns=[
    "error_code", "iso_name", "category", "severity", "is_customer_fault"
])
# Map to failure_category and failure_reason names
error_codes_df["failure_category"] = error_codes_df["category"]
error_codes_df["failure_reason"]   = error_codes_df["iso_name"]
print(f"    → {len(error_codes_df)} rows generated.")

# Build lookup maps
ec_map = error_codes_df.set_index("error_code")[["iso_name","category","severity","is_customer_fault"]].to_dict("index")

# Failure codes (all except "00")
FAILURE_CODES = [r[0] for r in ERROR_CODE_DATA if r[0] != "00"]
# Hardware codes
HW_CODES = [r[0] for r in ERROR_CODE_DATA if r[2] == "Hardware"]
# Non-hardware failure codes
NON_HW_FAILURE_CODES = [c for c in FAILURE_CODES if c not in HW_CODES]

# Weights for general failure code selection (uniform-ish, biased toward common ones)
BASE_FAILURE_WEIGHTS = {
    "51": 15, "55": 10, "57":  5, "61":  8, "65":  6,
    "91": 12, "96":  8, "68":  7, "05": 10, "14":  4,
    "41":  2, "43":  2, "54":  5, "62":  3, "76":  3,
    "85":  2, "34":  1, "13":  3, "30":  3, "31":  2,
    "44":  2, "52":  1, "53":  1, "58":  3, "75":  4,
    "P1":  5, "P2":  4, "P3":  3, "N1":  6,
}
fc_weights_arr = np.array([BASE_FAILURE_WEIGHTS[c] for c in FAILURE_CODES], dtype=float)
fc_weights_arr /= fc_weights_arr.sum()

# Hardware-elevated weights for NCR pattern
HW_FAILURE_WEIGHTS = fc_weights_arr.copy()
hw_indices = [FAILURE_CODES.index(c) for c in HW_CODES]
HW_FAILURE_WEIGHTS[hw_indices] *= 1.4
HW_FAILURE_WEIGHTS /= HW_FAILURE_WEIGHTS.sum()

# ═══════════════════════════════════════════════════════════════
# TABLE 3: transactions  (50,000 rows)
# ═══════════════════════════════════════════════════════════════
print("\n[3/4] Generating transactions (50,000 rows) — this may take a moment …")

N_TXN = 50_000
CHRONIC_OUTLIERS = {"ATM_015", "ATM_042", "ATM_087", "ATM_123", "ATM_156"}

# Terminal lookup for quick access
term_lookup = terminals_df.set_index("terminal_id")[["bank_owner","vendor","status"]].to_dict("index")
terminal_id_list = terminals_df["terminal_id"].tolist()

# ── timestamp generation (2024-01-01 to 2024-03-31) ──────────
TXN_START = datetime(2024, 1, 1)
TXN_END   = datetime(2024, 3, 31, 23, 59, 59)
total_seconds = int((TXN_END - TXN_START).total_seconds())

raw_seconds = np.random.randint(0, total_seconds + 1, N_TXN)
timestamps  = [TXN_START + timedelta(seconds=int(s)) for s in raw_seconds]

# Data quality issue: 1% of timestamps have seconds = 0
zero_sec_idx = np.random.choice(N_TXN, int(N_TXN * 0.01), replace=False)
timestamps = list(timestamps)
for idx in zero_sec_idx:
    ts = timestamps[idx]
    timestamps[idx] = ts.replace(second=0)

# ── terminal assignment (skewed so chronic outliers get enough txns) ──
# Give chronic outliers ~3x the probability
terminal_weights = np.ones(n_terminals)
for i, tid in enumerate(terminal_id_list):
    if tid in CHRONIC_OUTLIERS:
        terminal_weights[i] = 3.0
terminal_weights /= terminal_weights.sum()
terminal_assignments = np.random.choice(terminal_id_list, N_TXN, p=terminal_weights)

# ── transaction_type ──────────────────────────────────────────
txn_type_choices = ["Cash-Withdrawal", "Balance-Inquiry", "PIN-Change", "Mini-Statement"]
txn_type_probs   = [0.65, 0.20, 0.10, 0.05]
transaction_types = np.random.choice(txn_type_choices, N_TXN, p=txn_type_probs)

# ── amounts ───────────────────────────────────────────────────
amounts = np.where(
    transaction_types == "Cash-Withdrawal",
    (np.random.randint(1, 251, N_TXN) * 100),   # 100 to 25000 in multiples of 100
    0
)

# ── card_type & issuing_bank ──────────────────────────────────
card_types = np.where(np.random.random(N_TXN) < 0.60, "On-Us", "Off-Us")

issuing_banks = []
for i in range(N_TXN):
    tid  = terminal_assignments[i]
    own_bank = term_lookup[tid]["bank_owner"]
    if card_types[i] == "On-Us":
        issuing_banks.append(own_bank)
    else:
        other_banks = [b for b in BANKS if b != own_bank]
        issuing_banks.append(np.random.choice(other_banks))
issuing_banks = np.array(issuing_banks)

# ── PATTERN-AWARE status & error_code generation ─────────────
print("    Building failure patterns …")

statuses     = []
error_codes  = []

# Pre-compute is_weekend_peak mask
def is_weekend_peak(ts):
    dow  = ts.weekday()   # 5=Sat, 6=Sun
    hour = ts.hour
    if dow == 5 and 18 <= hour < 22:   # Saturday 6PM-10PM
        return True
    if dow == 6 and 10 <= hour < 14:   # Sunday 10AM-2PM
        return True
    return False

for i in range(N_TXN):
    tid    = terminal_assignments[i]
    ts     = timestamps[i]
    vendor = term_lookup[tid]["vendor"]
    ctype  = card_types[i]

    # Base failure probability = 12%
    p_fail = 0.12

    # Pattern 1: Chronic outlier terminals → 25-35% failure rate
    if tid in CHRONIC_OUTLIERS:
        p_fail = np.random.uniform(0.25, 0.35)

    # Pattern 2: Weekend peak hours → 2.5x multiplier (capped at 0.90)
    if is_weekend_peak(ts):
        p_fail = min(p_fail * 2.5, 0.90)

    # Pattern 4: Off-Us → 2.2x multiplier (capped at 0.90)
    if ctype == "Off-Us":
        p_fail = min(p_fail * 2.2, 0.90)

    # Determine status
    if np.random.random() < p_fail:
        statuses.append("Failed")
        # Pattern 3: NCR → higher HW failure rate
        if vendor == "NCR":
            ec = np.random.choice(FAILURE_CODES, p=HW_FAILURE_WEIGHTS)
        else:
            ec = np.random.choice(FAILURE_CODES, p=fc_weights_arr)
        error_codes.append(ec)
    else:
        statuses.append("Approved")
        error_codes.append("00")

statuses    = np.array(statuses)
error_codes = np.array(error_codes)

# ── response_time_ms ──────────────────────────────────────────
response_times = np.where(
    statuses == "Approved",
    np.random.normal(800,  200, N_TXN).clip(100,  3000),
    np.random.normal(2500, 800, N_TXN).clip(500,  8000),
).astype(float)

# Pattern 5: high latency → 85% failure (impose latency > 3500 on some failed txns)
high_latency_failed_idx = np.where(statuses == "Failed")[0]
n_high_latency = int(len(high_latency_failed_idx) * 0.30)
chosen = np.random.choice(high_latency_failed_idx, n_high_latency, replace=False)
response_times[chosen] = np.random.uniform(3501, 7000, n_high_latency)

# Data quality: 2% NULL response_time_ms
null_rt_idx = np.random.choice(N_TXN, int(N_TXN * 0.02), replace=False)
response_times[null_rt_idx] = np.nan
response_times = np.round(response_times, 1)

# ── failure_category & failure_reason from error_code ─────────
def get_ec_field(code, field):
    if code == "00":
        return None
    return ec_map.get(code, {}).get(field, None)

failure_categories = np.array([get_ec_field(c, "category") for c in error_codes])
failure_reasons    = np.array([get_ec_field(c, "iso_name")  for c in error_codes])

# ── Build transaction_ids ──────────────────────────────────────
transaction_ids = [f"TXN_{str(i).zfill(8)}" for i in range(1, N_TXN + 1)]

transactions_df = pd.DataFrame({
    "transaction_id":   transaction_ids,
    "terminal_id":      terminal_assignments,
    "timestamp":        [ts.strftime("%Y-%m-%d %H:%M:%S") for ts in timestamps],
    "transaction_type": transaction_types,
    "amount":           amounts,
    "status":           statuses,
    "error_code":       error_codes,
    "response_time_ms": response_times,
    "card_type":        card_types,
    "issuing_bank":     issuing_banks,
    "failure_category": failure_categories,
    "failure_reason":   failure_reasons,
})

# ── DATA QUALITY ISSUES ───────────────────────────────────────
print("    Injecting data quality issues …")

# (a) 3 exact duplicate transaction_id rows
dup_rows = transactions_df.iloc[[100, 500, 1000]].copy()
transactions_df = pd.concat([transactions_df, dup_rows], ignore_index=True)

# (b) 15 transactions with amount > 50000
extra_rows_idx = np.random.choice(len(transactions_df), 15, replace=False)
transactions_df.loc[transactions_df.index[extra_rows_idx], "amount"] = np.random.randint(50001, 100001, 15)

print(f"    → {len(transactions_df)} rows generated (incl. 3 duplicates).")

# ═══════════════════════════════════════════════════════════════
# TABLE 4: maintenance_logs  (~1500 rows)
# ═══════════════════════════════════════════════════════════════
print("\n[4/4] Generating maintenance_logs …")

N_MAINT = 1500
MAINT_START = datetime(2024, 1, 1)
MAINT_END   = datetime(2024, 4, 30, 23, 59, 59)
maint_total_sec = int((MAINT_END - MAINT_START).total_seconds())

COMPONENTS = ["Cash Cassette", "Card Reader", "Receipt Printer", "Keypad",
              "Dispenser Module", "Network Module", None]
ENGINEERS  = [f"ENG_{str(i).zfill(2)}" for i in range(1, 21)]
EVENT_TYPES = ["Preventive", "Corrective", "Emergency"]
EVENT_PROBS = [0.40, 0.45, 0.15]

RESOLUTION_TEMPLATES = {
    "Preventive":  [
        "Routine inspection completed. No issues found.",
        "Cash cassette cleaned and refilled. Software updated.",
        "Card reader cleaned. Firmware upgraded to latest version.",
        "Network settings verified. Response time within SLA.",
        "Preventive maintenance done. All components operational.",
    ],
    "Corrective":  [
        "Component replaced after failure detected during monitoring.",
        "Faulty unit removed and replaced. Tested post-installation.",
        "Intermittent failure resolved by replacing component.",
        "Error log cleared. Defective part replaced.",
        "Root cause identified as hardware wear. Replaced accordingly.",
    ],
    "Emergency":   [
        "Emergency callout — ATM was completely non-operational.",
        "Critical failure resolved on-site. ATM restored to service.",
        "Major hardware fault repaired under emergency protocol.",
        "Urgent replacement of failed component. Downtime minimised.",
        "Emergency maintenance done. Escalated to vendor for follow-up.",
    ],
}

# Skew terminal assignment: outlier ATMs get ~25% of maintenance
outlier_list    = list(CHRONIC_OUTLIERS)
non_outlier_ids = [t for t in terminal_id_list if t not in CHRONIC_OUTLIERS]

maint_term_weights = []
for tid in terminal_id_list:
    maint_term_weights.append(5.0 if tid in CHRONIC_OUTLIERS else 1.0)
maint_term_weights = np.array(maint_term_weights, dtype=float)
maint_term_weights /= maint_term_weights.sum()

maint_terminals   = np.random.choice(terminal_id_list, N_MAINT, p=maint_term_weights)
maint_event_types = np.random.choice(EVENT_TYPES, N_MAINT, p=EVENT_PROBS)
maint_secs        = np.random.randint(0, maint_total_sec + 1, N_MAINT)
maint_datetimes   = [
    (MAINT_START + timedelta(seconds=int(s))).strftime("%Y-%m-%d %H:%M:%S")
    for s in maint_secs
]

# Component replaced only for Corrective/Emergency
maint_components = []
for et in maint_event_types:
    if et in ["Corrective", "Emergency"]:
        maint_components.append(np.random.choice(COMPONENTS))
    else:
        maint_components.append(None)

maint_engineers = np.random.choice(ENGINEERS, N_MAINT)
maint_notes     = [
    np.random.choice(RESOLUTION_TEMPLATES[et]) for et in maint_event_types
]
maint_durations = np.round(np.random.uniform(0.5, 8.0, N_MAINT), 1)

# Cost based on event type
maint_costs = []
for et in maint_event_types:
    if et == "Preventive":
        maint_costs.append(int(np.random.randint(1000, 5001)))
    elif et == "Corrective":
        maint_costs.append(int(np.random.randint(5000, 25001)))
    else:  # Emergency
        maint_costs.append(int(np.random.randint(25000, 80001)))

maintenance_df = pd.DataFrame({
    "maintenance_id":     [f"MNT_{str(i).zfill(6)}" for i in range(1, N_MAINT + 1)],
    "terminal_id":        maint_terminals,
    "event_datetime":     maint_datetimes,
    "event_type":         maint_event_types,
    "component_replaced": maint_components,
    "engineer_id":        maint_engineers,
    "resolution_notes":   maint_notes,
    "duration_hours":     maint_durations,
    "cost_inr":           maint_costs,
})
print(f"    → {len(maintenance_df)} rows generated.")

# ═══════════════════════════════════════════════════════════════
# OUTPUT 1: CSVs
# ═══════════════════════════════════════════════════════════════
print("\n[Saving] Writing CSV files …")
terminals_df.to_csv(CSV_TERMS, index=False)
print(f"    ✓ {CSV_TERMS}")
error_codes_df.to_csv(CSV_ERRS, index=False)
print(f"    ✓ {CSV_ERRS}")
transactions_df.to_csv(CSV_TXN, index=False)
print(f"    ✓ {CSV_TXN}")
maintenance_df.to_csv(CSV_MAINT, index=False)
print(f"    ✓ {CSV_MAINT}")

# ═══════════════════════════════════════════════════════════════
# OUTPUT 2: Excel (one sheet per table)
# ═══════════════════════════════════════════════════════════════
print("\n[Saving] Writing Excel workbook …")
with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
    terminals_df.to_excel(writer,    sheet_name="atm_terminals",    index=False)
    error_codes_df.to_excel(writer,  sheet_name="error_codes",      index=False)
    transactions_df.to_excel(writer, sheet_name="transactions",     index=False)
    maintenance_df.to_excel(writer,  sheet_name="maintenance_logs", index=False)
print(f"    ✓ {EXCEL_PATH}")

# ═══════════════════════════════════════════════════════════════
# OUTPUT 3: SQLite database
# ═══════════════════════════════════════════════════════════════
print("\n[Saving] Writing SQLite database …")
conn = sqlite3.connect(DB_PATH)
cur  = conn.cursor()

# Drop existing tables/view
for obj in ["v_transaction_analysis", "transactions", "maintenance_logs", "error_codes", "atm_terminals"]:
    cur.execute(f"DROP VIEW IF EXISTS {obj}")
    cur.execute(f"DROP TABLE IF EXISTS {obj}")
conn.commit()

# Create tables with explicit DDL for proper typing
cur.executescript("""
CREATE TABLE IF NOT EXISTS atm_terminals (
    terminal_id    TEXT PRIMARY KEY,
    bank_owner     TEXT,
    vendor         TEXT,
    terminal_model TEXT,
    install_date   TEXT,
    location_name  TEXT,
    city           TEXT,
    state          TEXT,
    zone           TEXT,
    status         TEXT,
    latitude       REAL,
    longitude      REAL
);

CREATE TABLE IF NOT EXISTS error_codes (
    error_code        TEXT PRIMARY KEY,
    iso_name          TEXT,
    category          TEXT,
    severity          TEXT,
    is_customer_fault INTEGER,
    failure_category  TEXT,
    failure_reason    TEXT
);

CREATE TABLE IF NOT EXISTS transactions (
    transaction_id   TEXT,
    terminal_id      TEXT,
    timestamp        TEXT,
    transaction_type TEXT,
    amount           REAL,
    status           TEXT,
    error_code       TEXT,
    response_time_ms REAL,
    card_type        TEXT,
    issuing_bank     TEXT,
    failure_category TEXT,
    failure_reason   TEXT,
    FOREIGN KEY (terminal_id) REFERENCES atm_terminals(terminal_id),
    FOREIGN KEY (error_code)  REFERENCES error_codes(error_code)
);

CREATE TABLE IF NOT EXISTS maintenance_logs (
    maintenance_id      TEXT PRIMARY KEY,
    terminal_id         TEXT,
    event_datetime      TEXT,
    event_type          TEXT,
    component_replaced  TEXT,
    engineer_id         TEXT,
    resolution_notes    TEXT,
    duration_hours      REAL,
    cost_inr            INTEGER,
    FOREIGN KEY (terminal_id) REFERENCES atm_terminals(terminal_id)
);
""")
conn.commit()

# Load data
terminals_df.to_sql("atm_terminals",    conn, if_exists="append", index=False)
error_codes_df.to_sql("error_codes",    conn, if_exists="append", index=False)
transactions_df.to_sql("transactions",  conn, if_exists="append", index=False)
maintenance_df.to_sql("maintenance_logs", conn, if_exists="append", index=False)
print("    ✓ Tables loaded into SQLite.")

# Create view
VIEW_DDL = """
CREATE VIEW IF NOT EXISTS v_transaction_analysis AS
SELECT 
  t.transaction_id, t.terminal_id, t.timestamp,
  strftime('%Y-%m-%d', t.timestamp) as txn_date,
  strftime('%H', t.timestamp) as txn_hour,
  strftime('%w', t.timestamp) as day_of_week,
  t.transaction_type, t.amount, t.status,
  t.error_code, t.response_time_ms, t.card_type, t.issuing_bank,
  t.failure_category, t.failure_reason,
  ec.iso_name, ec.severity, ec.is_customer_fault,
  a.bank_owner, a.vendor, a.terminal_model, a.install_date,
  a.location_name, a.city, a.state, a.zone,
  a.status as terminal_status, a.latitude, a.longitude,
  CASE WHEN t.status = 'Failed' THEN 1 ELSE 0 END as is_failed,
  CASE WHEN t.response_time_ms > 3000 THEN 'High-Latency' ELSE 'Normal' END as latency_band,
  CASE WHEN t.amount > 10000 THEN 'High-Value' 
       WHEN t.amount > 2000  THEN 'Medium-Value' 
       ELSE 'Low-Value' END as value_band
FROM transactions t
LEFT JOIN atm_terminals a ON t.terminal_id = a.terminal_id
LEFT JOIN error_codes ec  ON t.error_code  = ec.error_code;
"""
cur.execute(VIEW_DDL)
conn.commit()
print(f"    ✓ View v_transaction_analysis created.")
print(f"    ✓ {DB_PATH}")

# ═══════════════════════════════════════════════════════════════
# VALIDATION REPORT
# ═══════════════════════════════════════════════════════════════
print("\n" + "=" * 65)
print("  VALIDATION REPORT")
print("=" * 65)

# Row counts
print("\n── Row Counts ──────────────────────────────────────────────")
rc = {
    "atm_terminals":    cur.execute("SELECT COUNT(*) FROM atm_terminals").fetchone()[0],
    "error_codes":      cur.execute("SELECT COUNT(*) FROM error_codes").fetchone()[0],
    "transactions":     cur.execute("SELECT COUNT(*) FROM transactions").fetchone()[0],
    "maintenance_logs": cur.execute("SELECT COUNT(*) FROM maintenance_logs").fetchone()[0],
    "view_rows":        cur.execute("SELECT COUNT(*) FROM v_transaction_analysis").fetchone()[0],
}
for name, count in rc.items():
    print(f"    {name:<25} {count:>7,} rows")

# FK integrity
print("\n── FK Integrity ────────────────────────────────────────────")
bad_fk_term = cur.execute("""
    SELECT COUNT(*) FROM transactions t
    WHERE t.terminal_id NOT IN (SELECT terminal_id FROM atm_terminals)
""").fetchone()[0]
bad_fk_ec = cur.execute("""
    SELECT COUNT(*) FROM transactions t
    WHERE t.error_code NOT IN (SELECT error_code FROM error_codes)
""").fetchone()[0]
bad_fk_maint = cur.execute("""
    SELECT COUNT(*) FROM maintenance_logs m
    WHERE m.terminal_id NOT IN (SELECT terminal_id FROM atm_terminals)
""").fetchone()[0]
print(f"    transactions → atm_terminals FK violations : {bad_fk_term}")
print(f"    transactions → error_codes FK violations   : {bad_fk_ec}")
print(f"    maintenance  → atm_terminals FK violations : {bad_fk_maint}")
fk_ok = (bad_fk_term == 0 and bad_fk_ec == 0 and bad_fk_maint == 0)
print(f"    FK integrity: {'✓ 100% VALID' if fk_ok else '✗ ISSUES FOUND'}")

# Overall failure rate
print("\n── Overall Failure Rate ────────────────────────────────────")
tot   = cur.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
fail  = cur.execute("SELECT COUNT(*) FROM transactions WHERE status='Failed'").fetchone()[0]
print(f"    Total transactions : {tot:>7,}")
print(f"    Failed             : {fail:>7,}  ({fail/tot*100:.2f}%)")
print(f"    Approved           : {tot-fail:>7,}  ({(tot-fail)/tot*100:.2f}%)")

# Data quality checks
print("\n── Data Quality Issues (Embedded) ──────────────────────────")
null_rt = cur.execute("SELECT COUNT(*) FROM transactions WHERE response_time_ms IS NULL").fetchone()[0]
dups    = cur.execute("""
    SELECT COUNT(*) FROM (
        SELECT transaction_id FROM transactions
        GROUP BY transaction_id HAVING COUNT(*) > 1
    )
""").fetchone()[0]
outlier_amt = cur.execute("SELECT COUNT(*) FROM transactions WHERE amount > 50000").fetchone()[0]
zero_sec    = cur.execute("""
    SELECT COUNT(*) FROM transactions
    WHERE strftime('%S', timestamp) = '00'
""").fetchone()[0]
print(f"    NULL response_time_ms          : {null_rt:>5,}  (target ~{int(tot*0.02):,})")
print(f"    Duplicate transaction_id groups: {dups:>5,}  (target  3)")
print(f"    Amount > 50,000 outliers       : {outlier_amt:>5,}  (target 15)")
print(f"    Timestamps with seconds=00     : {zero_sec:>5,}  (target ~{int(50000*0.01):,})")

# Pattern spot-checks
print("\n── Pattern Verification ────────────────────────────────────")

# Pattern 1: Chronic outliers
print("\n  Pattern 1 — Chronic Outlier Terminals (target 25-35% fail):")
for atm in sorted(CHRONIC_OUTLIERS):
    r = cur.execute("""
        SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
        FROM transactions WHERE terminal_id=?
    """, (atm,)).fetchone()
    n, f = r
    pct = (f/n*100) if n else 0
    print(f"    {atm}: {n:,} txns, {f:,} failed ({pct:.1f}%)")

# Pattern 2: Weekend peak hours
print("\n  Pattern 2 — Weekend Peak Hours (target ~2.5x higher failure):")
# Saturday 6-10PM = dow 6 in SQLite (strftime %w: 0=Sunday,6=Saturday)
sat_peak = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions
    WHERE strftime('%w', timestamp)='6'
      AND CAST(strftime('%H', timestamp) AS INTEGER) BETWEEN 18 AND 21
""").fetchone()
sun_peak = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions
    WHERE strftime('%w', timestamp)='0'
      AND CAST(strftime('%H', timestamp) AS INTEGER) BETWEEN 10 AND 13
""").fetchone()
off_peak = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions
    WHERE NOT (
        (strftime('%w', timestamp)='6' AND CAST(strftime('%H', timestamp) AS INTEGER) BETWEEN 18 AND 21)
     OR (strftime('%w', timestamp)='0' AND CAST(strftime('%H', timestamp) AS INTEGER) BETWEEN 10 AND 13)
    )
""").fetchone()
def pct(r): return (r[1]/r[0]*100) if r[0] else 0
print(f"    Saturday 6-10PM : {sat_peak[0]:,} txns, fail rate {pct(sat_peak):.1f}%")
print(f"    Sunday 10AM-2PM : {sun_peak[0]:,} txns, fail rate {pct(sun_peak):.1f}%")
print(f"    Off-peak        : {off_peak[0]:,} txns, fail rate {pct(off_peak):.1f}%")

# Pattern 3: NCR hardware failure rate
print("\n  Pattern 3 — NCR vs Others Hardware Failure Rate:")
ncr_hw = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN t.failure_category='Hardware' THEN 1 ELSE 0 END)
    FROM transactions t
    JOIN atm_terminals a ON t.terminal_id=a.terminal_id
    WHERE a.vendor='NCR' AND t.status='Failed'
""").fetchone()
other_hw = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN t.failure_category='Hardware' THEN 1 ELSE 0 END)
    FROM transactions t
    JOIN atm_terminals a ON t.terminal_id=a.terminal_id
    WHERE a.vendor<>'NCR' AND t.status='Failed'
""").fetchone()
ncr_hw_pct   = (ncr_hw[1]/ncr_hw[0]*100)   if ncr_hw[0]   else 0
other_hw_pct = (other_hw[1]/other_hw[0]*100) if other_hw[0] else 0
print(f"    NCR terminals   : {ncr_hw[0]:,} failures, HW failures {ncr_hw[1]:,} ({ncr_hw_pct:.1f}%)")
print(f"    Other vendors   : {other_hw[0]:,} failures, HW failures {other_hw[1]:,} ({other_hw_pct:.1f}%)")

# Pattern 4: Off-Us vs On-Us failure rate
print("\n  Pattern 4 — Off-Us vs On-Us Failure Rate (target 2.2x):")
onus   = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions WHERE card_type='On-Us'
""").fetchone()
offus  = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions WHERE card_type='Off-Us'
""").fetchone()
onus_pct  = pct(onus)
offus_pct = pct(offus)
ratio = (offus_pct / onus_pct) if onus_pct else 0
print(f"    On-Us  : {onus[0]:,} txns, fail rate {onus_pct:.1f}%")
print(f"    Off-Us : {offus[0]:,} txns, fail rate {offus_pct:.1f}%")
print(f"    Ratio  : {ratio:.2f}x  (target ~2.2x)")

# Pattern 5: High latency → failure rate
print("\n  Pattern 5 — High Latency (>3500ms) → Failure Rate (target ~85%):")
high_lat = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions WHERE response_time_ms > 3500
""").fetchone()
norm_lat = cur.execute("""
    SELECT COUNT(*), SUM(CASE WHEN status='Failed' THEN 1 ELSE 0 END)
    FROM transactions WHERE response_time_ms <= 3500 AND response_time_ms IS NOT NULL
""").fetchone()
print(f"    response_time_ms > 3500 : {high_lat[0]:,} txns, fail rate {pct(high_lat):.1f}%  (target ≥85%)")
print(f"    Normal latency          : {norm_lat[0]:,} txns, fail rate {pct(norm_lat):.1f}%")

conn.close()
print("\n" + "=" * 65)
print("  ✓ All outputs generated successfully!")
print("=" * 65)
print(f"\nOutput files:")
print(f"  Excel   → {EXCEL_PATH}")
print(f"  CSVs    → {DATA_DIR}/")
print(f"  SQLite  → {DB_PATH}")
