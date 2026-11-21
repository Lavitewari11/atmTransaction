# p096: ATM Transaction Failure Analysis — Build Plan (BUILD_PLAN.md)

## 1. Project Summary & Deliverables Index
- **Project Code & Title:** `p096` — ATM Transaction Failure Analysis
- **Domain:** Banking & Payment Network Operational Data Analytics
- **Candidate:** Lavi Tewari (Roll No: 2500320140085), MCA 2nd Year, ABES Institute of Technology
- **Academic Mentor:** Mr. Dinesh Sereno
- **Primary Objective:** Build an interactive multi-page Streamlit analytics application diagnosing ATM failures across 50,003 multi-bank records.
- **Key Capabilities:** Standardized 4-tier banking failure taxonomy, rolling 7-day SLA outlier terminal detection (>5.0%), 24x7 temporal congestion heatmap, MTBF predictive maintenance calculations, and network latency correlation.
- **Target Tech Stack:** Python 3.12, Streamlit, Pandas, NumPy, Plotly, SQLite, pytest, Streamlit Community Cloud.

### Project Deliverables Index
- Step 01: [docs/01_Project_Understanding.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/01_Project_Understanding.docx)
- Step 02: [docs/02_HMW_Statements.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/02_HMW_Statements.docx)
- Step 03: [docs/03_Needs_Analysis.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/03_Needs_Analysis.docx)
- Step 04: [docs/04_Traceability_Matrix.xlsx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/04_Traceability_Matrix.xlsx)
- Step 05: [docs/05_BRD.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/05_BRD.docx)
- Step 07: [docs/07_TRD.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/07_TRD.docx)
- Step 08: [docs/08_Database_Design.xlsx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/08_Database_Design.xlsx)
- Step 09: [docs/09_ER_Diagram.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/09_ER_Diagram.docx)
- Step 10: [data/app.db](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/data/app.db) & [src/generate_data.py](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/src/generate_data.py)
- Step 11: [docs/11_User_Stories.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/11_User_Stories.docx)
- Step 12: [docs/12_UI_Framework.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/12_UI_Framework.docx) & [src/ui.py](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/src/ui.py)
- Step 13: [docs/13_Test_Cases.docx](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/docs/13_Test_Cases.docx) & [tests/](file:///c:/Users/ADMIN/OneDrive/Desktop/AtmTransactionAnalysis/tests)

---

## 2. Final Repository Folder Structure
```text
atm-transaction-analysis/
├── app.py                         # Main entry point & session state router
├── pages/
│   ├── 01_Fleet_Overview.py       # OQ-01, US-04, US-06 (KPI row + hourly trend line)
│   ├── 02_Failure_Category.py     # OQ-02, US-07, US-09 (Pareto 80/20 + type breakdown)
│   ├── 03_Terminal_Health.py      # OQ-03, US-10, US-11, US-12 (Outliers >5% + MTBF)
│   ├── 04_Temporal_Patterns.py    # OQ-04, US-08 (24x7 congestion heatmap)
│   ├── 05_Network_Latency.py      # OQ-06, US-13, US-14 (Scatter + On/Off-Us Z-test)
│   └── 06_Financial_Impact.py     # OQ-08, US-06 (INR at risk treemap & cards)
├── src/
│   ├── db.py                      # SQLite connectors & @st.cache_data loaders
│   ├── transform.py               # Taxonomy mapping, cleaning & quality metrics
│   ├── analysis.py                # Aggregation, outlier detection, MTBF, stats
│   ├── charts.py                  # Standardized Plotly chart factory
│   ├── ui.py                      # Tailwind header, KPI cards, sidebar filters
│   └── generate_data.py           # Synthetic dataset generator (50k rows)
├── data/
│   ├── app.db                     # SQLite database with 4 tables + 1 view
│   ├── atm_terminals.csv
│   ├── error_codes.csv
│   ├── transactions.csv
│   ├── maintenance_logs.csv
│   └── atm_transactions_data.xlsx # Complete 4-sheet master workbook
├── tests/
│   ├── conftest.py
│   ├── test_data.py               # TC-T-01..TC-T-05
│   ├── test_analysis.py           # TC-T-06..TC-T-09
│   └── test_app.py                # TC-T-10..TC-T-11
├── docs/                          # All Word and Excel deliverables
├── .streamlit/
│   └── config.toml                # Navy/Slate UI theme configuration
├── requirements.txt
└── README.md
```

---

## 3. Step-by-Step Build Tasks

### BUILD-01: Data Ingestion & Caching Layer (`src/db.py`)
- **Goal:** Connect to SQLite `data/app.db` in read-only mode, cache master and fact DataFrames using `@st.cache_data`.
- **Linked Stories:** US-01 | **Features:** FR-01, FR-02
- **Pass Criteria:** `pytest tests/test_data.py` passes 100%.

### BUILD-02: Data Transformation & Hygiene Pipeline (`src/transform.py`)
- **Goal:** Enrich raw transactions with failure taxonomy, detect quarantined rows (missing latency, duplicates, outliers).
- **Linked Stories:** US-02, US-03 | **Features:** FR-02, FR-03
- **Pass Criteria:** 0 unhandled NULL categories; quality telemetry dict contains exact counts.

### BUILD-03: Core Statistical & Aggregation Engine (`src/analysis.py`)
- **Goal:** Implement failure rate aggregations, outlier detection (>5% SLA), MTBF calculations, and Pearson correlation.
- **Linked Stories:** US-04, US-10, US-11, US-13 | **Features:** FR-04, FR-05, FR-08, FR-09, FR-13
- **Pass Criteria:** `pytest tests/test_analysis.py` passes 100%.

### BUILD-04: Centralized Plotly Chart Factory (`src/charts.py`)
- **Goal:** Build reusable chart generators (trend line, Pareto bar+line, outlier H-bar, 24x7 heatmap, latency scatter, donut).
- **Linked Stories:** US-04, US-07, US-08, US-10, US-12, US-13, US-14 | **Features:** FR-14
- **Pass Criteria:** All figure generators return valid Plotly Graph Objects without warnings.

### BUILD-05: Application Entrypoint & Sidebar Filter Engine (`app.py`)
- **Goal:** Setup page router, render global Tailwind top header, and synchronize sidebar filter state.
- **Linked Stories:** US-01, US-05, US-15 | **Features:** FR-15, FR-16
- **Pass Criteria:** App launches with `streamlit run app.py` without console errors.

### BUILD-06: Multipage Implementations (`pages/01` to `pages/06`)
- **Goal:** Build the 6 specialized analytical views integrating KPI rows, Plotly charts, and diagnostic insight callouts.
- **Linked Stories:** US-04 through US-14 | **Features:** FR-04 through FR-14
- **Pass Criteria:** All 6 screens navigate smoothly, respond to sidebar filters in < 1.5s.

---

## 4. Definition of Done (DoD)
1. 100% of Must-Have and Should-Have Business Requirements (BR-01..BR-10) fully functional.
2. All 11 pytest unit tests pass with zero errors.
3. Streamlit app runs locally with zero traceback exceptions.
4. SQLite database file `data/app.db` committed with valid foreign keys.
5. Pinned `requirements.txt` validated in fresh virtual environment.

---

## 5. Execution Commands
```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run unit tests
pytest -v

# 3. Launch Streamlit analytics dashboard
streamlit run app.py
```

---

## 6. Viva Preparation: 10 Critical Questions & Model Answers
1. **Q: Why was SQLite selected over PostgreSQL/MySQL for this project?**  
   *A: SQLite provides zero-configuration, self-contained relational storage that runs in-process with Python. It eliminates external server dependencies, ensures deterministic grading, and deploys directly within Streamlit Community Cloud's 1 GB memory limit.*
2. **Q: How does `@st.cache_data` enhance application responsiveness?**  
   *A: It serializes function outputs to memory. Once 50,003 records are read from SQLite, subsequent filter changes operate on memory copies without disk I/O, cutting dashboard reload latency from seconds to < 200 ms.*
3. **Q: What statistical formula detects chronic terminal outliers?**  
   *A: A dual-condition filter: rolling 7-day failure rate exceeding the 5.0% banking SLA threshold, coupled with Z-score exceeding $\mu + 2\sigma$ relative to the fleet mean.*
4. **Q: How is the 24x7 heatmap calculated?**  
   *A: Using a 2D pivot table grouping transaction counts by `timestamp.dt.hour` (0..23) and `timestamp.dt.day_name()` (Monday..Sunday), rendered via Plotly `density_heatmap`.*
5. **Q: How was the 2.2x failure disparity between On-Us and Off-Us transactions statistically proven?**  
   *A: By performing a two-proportion Z-test comparing failed transaction ratios across On-Us and Off-Us volumes, yielding a p-value < 0.001 indicating statistically significant interbank switch degradation.*
6. **Q: What is the Mean Time Between Failures (MTBF) formula?**  
   *A: $MTBF = \frac{\text{Total Operating Hours} - \text{Cumulative Downtime}}{\text{Total Failure Count}}$. It quantifies machine reliability in operating hours.*
7. **Q: How does the application handle data quality anomalies?**  
   *A: Missing response times are imputed or flagged, duplicate transaction IDs are quarantined from aggregations, and extreme amounts (> ₹50,000) are audited in a dedicated sidebar badge.*
8. **Q: Why use Tailwind CSS with Streamlit?**  
   *A: Tailwind provides modern enterprise UI components (custom colored borders, flexbox KPI rows, status badges) that elevate the application above generic Streamlit prototypes.*
9. **Q: How does the Pareto chart operationalize the 80/20 rule?**  
   *A: It sorts failure categories descending on a bar chart and overlays a cumulative percentage curve. Operators immediately see that Hardware and Network-Switch issues cause > 75% of all outages.*
10. **Q: How does this project prevent Scope Creep?**  
    *A: Every feature is traced through the RTM (OQ -> HMW -> BR -> US -> Screen -> Test). Any feature lacking a verified BR was explicitly designated Out-of-Scope.*
