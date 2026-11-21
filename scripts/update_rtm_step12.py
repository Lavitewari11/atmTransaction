import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.load_workbook("docs/04_Traceability_Matrix.xlsx")
ws = wb["Traceability_Matrix"]

# Update rows with DB Table.Column, User Story, and UI Screen mappings
# Row structure:
# Row 4: BR-01
# Row 5: BR-02
# Row 6: BR-03
# Row 7: BR-04
# Row 8: BR-05
# Row 9: BR-06
# Row 10: BR-07
# Row 11: BR-08
# Row 12: BR-09
# Row 13: BR-10
# Row 14: BR-11
# Row 15: BR-12

mappings = {
    4: {  # BR-01
        "db": "transactions.timestamp, status, terminal_id",
        "us": "US-01, US-04, US-06",
        "ui": "UI-01 (Fleet Overview)"
    },
    5: {  # BR-02
        "db": "error_codes.error_code, failure_category",
        "us": "US-02, US-03, US-07",
        "ui": "UI-02 (Failure Category)"
    },
    6: {  # BR-03
        "db": "transactions.terminal_id, status, atm_terminals.location",
        "us": "US-10",
        "ui": "UI-03 (Terminal Health)"
    },
    7: {  # BR-04
        "db": "transactions.timestamp, failure_category",
        "us": "US-08",
        "ui": "UI-04 (Temporal Patterns)"
    },
    8: {  # BR-05
        "db": "transactions.transaction_type, status, amount",
        "us": "US-07, US-09",
        "ui": "UI-02 (Failure Category)"
    },
    9: {  # BR-06
        "db": "transactions.response_time_ms, status",
        "us": "US-13",
        "ui": "UI-05 (Network Latency)"
    },
    10: { # BR-07
        "db": "transactions.card_type, issuing_bank, status",
        "us": "US-14",
        "ui": "UI-05 (Network Latency)"
    },
    11: { # BR-08
        "db": "transactions.amount, status, failure_category",
        "us": "US-06",
        "ui": "UI-06 (Financial Impact)"
    },
    12: { # BR-09
        "db": "atm_terminals.vendor, terminal_model, status",
        "us": "US-12",
        "ui": "UI-03 (Terminal Health)"
    },
    13: { # BR-10
        "db": "transactions.timestamp, maintenance_logs.event_datetime",
        "us": "US-11",
        "ui": "UI-03 (Terminal Health)"
    },
    14: { # BR-11
        "db": "transactions.*, atm_terminals.*",
        "us": "US-01, US-05",
        "ui": "UI-01..UI-06 (Global Sidebar)"
    },
    15: { # BR-12
        "db": "v_transaction_analysis.*",
        "us": "US-15",
        "ui": "UI-08 (Data Export)"
    }
}

for row_idx, data in mappings.items():
    ws.cell(row=row_idx, column=9, value=data["db"])   # DB table.column
    ws.cell(row=row_idx, column=10, value=data["us"])  # User Story
    ws.cell(row=row_idx, column=11, value=data["ui"])  # UI Screen
    # Update status to Designed
    ws.cell(row=row_idx, column=7, value="Designed")

wb.save("docs/04_Traceability_Matrix.xlsx")
print("RTM updated with DB, US, and UI mappings.")
