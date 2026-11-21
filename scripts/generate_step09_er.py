import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_shd(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_margins(cell, top=100, bottom=100, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, v in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        n = OxmlElement(f'w:{m}')
        n.set(qn('w:w'), str(v))
        n.set(qn('w:type'), 'dxa')
        tcMar.append(n)
    tcPr.append(tcMar)

def stylize(table, widths, hbg="1A365D", alt="F7FAFC"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.rows[0]._tr.get_or_add_trPr().append(
        parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for c in table.rows[0].cells:
        set_shd(c, hbg); set_margins(c, 110, 110, 120, 120)
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.bold = True; r.font.color.rgb = RGBColor(255,255,255)
                r.font.size = Pt(9.5); r.font.name = "Calibri"
    for ri, row in enumerate(table.rows[1:], 1):
        bg = alt if ri % 2 else "FFFFFF"
        for c in row.cells:
            set_shd(c, bg); set_margins(c, 90, 90, 110, 110)
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9); r.font.name = "Calibri"
                    r.font.color.rgb = RGBColor(45,55,72)
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)

def callout(doc, text, title="NOTE", bg="EDF2F7", border="2B6CB0"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0,0)
    set_shd(cell, bg); set_margins(cell, 130, 130, 160, 160)
    cell.width = Inches(6.5)
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'''<w:tcBorders {nsdecls("w")}>
        <w:left w:val="single" w:sz="24" w:space="0" w:color="{border}"/>
        <w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>
    </w:tcBorders>'''))
    p = cell.paragraphs[0]
    rt = p.add_run(f"[{title}]  ")
    rt.bold = True; rt.font.name="Calibri"; rt.font.size=Pt(9.5)
    rt.font.color.rgb=RGBColor(43,108,176)
    rb = p.add_run(text)
    rb.font.name="Calibri"; rb.font.size=Pt(9.5)
    rb.font.color.rgb=RGBColor(45,55,72)
    doc.add_paragraph()

def head(doc, text, lv=1, color=(26,54,93)):
    h = doc.add_heading(level=lv)
    r = h.add_run(text)
    r.font.color.rgb = RGBColor(*color)
    r.font.name = "Calibri"

def fill_table(table, data):
    for ri, row in enumerate(data, 1):
        for ci, txt in enumerate(row):
            table.rows[ri].cells[ci].paragraphs[0].text = txt

def build_er_doc(out):
    doc = Document()
    for s in doc.sections:
        s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Inches(1.0)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("p096: ATM TRANSACTION FAILURE ANALYSIS")
    r.bold=True; r.font.size=Pt(21); r.font.name="Calibri"
    r.font.color.rgb=RGBColor(26,54,93)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("Step 9 — Entity-Relationship Diagram Specification  v1.0")
    r2.font.size=Pt(13); r2.font.name="Calibri"
    r2.font.color.rgb=RGBColor(74,85,104)

    # Meta
    meta = doc.add_table(rows=5, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    mdata = [
        ("Project & Code",   "p096 | ATM Transaction Failure Analysis"),
        ("Author / Student", "Lavi Tewari  |  Roll: 2500320140085  |  MCA 2nd Yr DA, ABES"),
        ("Mentor",           "Mr. Dinesh Sereno"),
        ("Source Document",  "Step 8 Database Design (08_Database_Design.xlsx)"),
        ("Date & Status",    "November 2026 | Baselined v1.0"),
    ]
    for i,(k,v) in enumerate(mdata):
        c0,c1 = meta.rows[i].cells
        set_shd(c0,"F0F4F8"); set_shd(c1,"FFFFFF")
        set_margins(c0,60,60,100,100); set_margins(c1,60,60,100,100)
        r0=c0.paragraphs[0].add_run(k); r0.bold=True
        r0.font.size=Pt(9.5); r0.font.name="Calibri"; r0.font.color.rgb=RGBColor(26,54,93)
        r1=c1.paragraphs[0].add_run(v)
        r1.font.size=Pt(9.5); r1.font.name="Calibri"; r1.font.color.rgb=RGBColor(45,55,72)
        c0.width=Inches(2.2); c1.width=Inches(4.3)
    doc.add_paragraph()

    # 1. Mermaid Code
    head(doc,"1. Mermaid ER Diagram — Source Code")
    doc.add_paragraph("Paste the Mermaid code below at mermaid.live → export PNG → insert into Section 2 of this document.")

    mermaid_code = '''erDiagram
    ATM_TERMINALS {
        TEXT terminal_id PK "PK NOT NULL"
        TEXT bank_owner "NOT NULL"
        TEXT vendor "NOT NULL"
        TEXT terminal_model "NOT NULL"
        DATE install_date "NOT NULL"
        TEXT location_name "NOT NULL"
        TEXT city "NOT NULL"
        TEXT state "NOT NULL"
        TEXT zone "Metro|Urban|Semi-Urban"
        TEXT status "Active|Inactive"
        REAL latitude "nullable"
        REAL longitude "nullable"
    }

    ERROR_CODES {
        TEXT error_code PK "PK NOT NULL"
        TEXT iso_name "NOT NULL"
        TEXT description "NOT NULL"
        TEXT failure_category "Hardware|Network-Switch|Customer-Account|Host-Timeout"
        TEXT severity "Critical|High|Medium|Low"
        INTEGER is_customer_fault "0|1 NOT NULL"
    }

    TRANSACTIONS {
        TEXT transaction_id PK "PK NOT NULL"
        TEXT terminal_id FK "FK -> ATM_TERMINALS NOT NULL"
        DATETIME timestamp "NOT NULL"
        TEXT transaction_type "Cash-Withdrawal|Balance-Inquiry|PIN-Change|Mini-Statement"
        REAL amount "CHECK >= 0"
        TEXT status "Approved|Failed|Reversed"
        TEXT error_code FK "FK -> ERROR_CODES nullable"
        INTEGER response_time_ms "CHECK > 0"
        TEXT card_type "On-Us|Off-Us"
        TEXT issuing_bank "NOT NULL"
        TEXT failure_category "nullable"
        TEXT failure_reason "nullable"
    }

    MAINTENANCE_LOGS {
        TEXT maintenance_id PK "PK NOT NULL"
        TEXT terminal_id FK "FK -> ATM_TERMINALS NOT NULL"
        DATETIME event_datetime "NOT NULL"
        TEXT event_type "Preventive|Corrective|Emergency"
        TEXT component_replaced "nullable"
        TEXT engineer_id "NOT NULL"
        TEXT resolution_notes "nullable"
        REAL duration_hours "CHECK > 0"
        REAL cost_inr "CHECK >= 0"
    }

    ATM_TERMINALS ||--o{ TRANSACTIONS : "processes"
    ATM_TERMINALS ||--o{ MAINTENANCE_LOGS : "receives"
    ERROR_CODES ||--o{ TRANSACTIONS : "classifies"'''

    p_code = doc.add_paragraph()
    r_code = p_code.add_run(mermaid_code)
    r_code.font.name="Courier New"; r_code.font.size=Pt(8)
    r_code.font.color.rgb=RGBColor(26,54,93)
    doc.add_paragraph()

    callout(doc,
        "To render the PNG: 1) Visit https://mermaid.live  2) Paste the code above "
        "3) Click 'Actions' → 'PNG' → Download  4) Insert the downloaded PNG in Section 2 below.",
        title="EXPORT INSTRUCTIONS")

    # 2. PNG Placeholder
    head(doc,"2. ER Diagram PNG (Insert Exported Image Here)")
    p_placeholder = doc.add_paragraph()
    p_placeholder.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ph = p_placeholder.add_run(
        "[PLACEHOLDER — Paste the mermaid.live exported PNG here]\n"
        "File: docs/09_ER_Diagram.png"
    )
    r_ph.bold=True; r_ph.font.name="Calibri"; r_ph.font.size=Pt(11)
    r_ph.font.color.rgb=RGBColor(74,85,104)
    
    # Draw a simple placeholder box
    ph_tbl = doc.add_table(rows=1, cols=1)
    ph_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    ph_cell = ph_tbl.cell(0,0)
    set_shd(ph_cell,"F0F4F8")
    set_margins(ph_cell,200,200,200,200)
    ph_cell.width = Inches(6.0)
    tcPr = ph_cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'''<w:tcBorders {nsdecls("w")}>
        <w:left w:val="single" w:sz="12" w:space="0" w:color="2B6CB0"/>
        <w:top w:val="single" w:sz="12" w:space="0" w:color="2B6CB0"/>
        <w:right w:val="single" w:sz="12" w:space="0" w:color="2B6CB0"/>
        <w:bottom w:val="single" w:sz="12" w:space="0" w:color="2B6CB0"/>
    </w:tcBorders>'''))
    for _ in range(6):
        ph_cell.add_paragraph()
    ph_cell.paragraphs[3].alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_img = ph_cell.paragraphs[3].add_run("[ ER Diagram PNG — Insert Here ]")
    r_img.bold=True; r_img.font.name="Calibri"; r_img.font.size=Pt(12)
    r_img.font.color.rgb=RGBColor(43,108,176)
    for _ in range(6):
        ph_cell.add_paragraph()
    doc.add_paragraph()

    # 3. Plain English Relationships
    head(doc,"3. Entity Relationships — Plain-English Description")
    rel_tbl = doc.add_table(rows=4, cols=4)
    rel_tbl.rows[0].cells[0].paragraphs[0].text = "Relationship"
    rel_tbl.rows[0].cells[1].paragraphs[0].text = "Cardinality"
    rel_tbl.rows[0].cells[2].paragraphs[0].text = "Verb Phrase"
    rel_tbl.rows[0].cells[3].paragraphs[0].text = "Plain-English Statement"
    rels = [
        ("atm_terminals → transactions","One-to-Many (1 : 0..∞)","processes",
         "One ATM terminal processes zero or many transactions over its operational lifetime. Each transaction must belong to exactly one terminal."),
        ("atm_terminals → maintenance_logs","One-to-Many (1 : 0..∞)","receives",
         "One ATM terminal receives zero or many scheduled or emergency maintenance events. Each maintenance event must reference exactly one terminal."),
        ("error_codes → transactions","One-to-Many (1 : 0..∞)","classifies",
         "One error code classifies zero or many failed transactions. Each failed transaction carries exactly one error code; approved transactions carry error code '00'."),
    ]
    fill_table(rel_tbl, rels)
    stylize(rel_tbl,[1.5,1.3,0.9,2.8])
    doc.add_paragraph()

    # 4. Business Rules Enforced
    head(doc,"4. Business Rules Enforced by the Data Model")
    biz_rules = [
        ("DBR-01","Referential Integrity — Terminal","Every transactions.terminal_id and maintenance_logs.terminal_id MUST exist in atm_terminals.terminal_id. Enforced via FOREIGN KEY REFERENCES in SQLite with PRAGMA foreign_keys = ON."),
        ("DBR-02","Referential Integrity — Error Code","Every non-NULL transactions.error_code MUST exist in error_codes.error_code. Approved transactions must carry error_code = '00'."),
        ("DBR-03","Transaction Amount Constraint","amount >= 0 enforced by CHECK. Cash-Withdrawal amount must be multiples of ₹100; Balance-Inquiry, PIN-Change, and Mini-Statement amount = 0."),
        ("DBR-04","Status-Error Code Consistency","If status = 'Approved', error_code = '00'. If status = 'Failed', error_code ≠ '00' and failure_category IS NOT NULL. This is enforced at application-layer in src/transform.py."),
        ("DBR-05","Response Time Validity","response_time_ms > 0 enforced by CHECK. NULL values are flagged as data quality issues and recorded in data quality log."),
        ("DBR-06","Zone Classification Consistency","zone must be one of Metro | Urban | Semi-Urban. Metro terminals are in Tier-1 cities (Mumbai, Delhi, Bengaluru, Chennai, Kolkata, Hyderabad). Enforced by CHECK constraint."),
        ("DBR-07","Maintenance Cost Bounds","cost_inr >= 0 enforced by CHECK. Emergency events minimum cost = ₹25,000; Preventive maximum cost = ₹5,000. Enforced at generator level."),
        ("DBR-08","Unique Transaction ID","transaction_id is PRIMARY KEY — uniqueness guaranteed at DB level. The 3 intentional duplicate rows are detected and quarantined during data cleaning."),
    ]
    rules_tbl = doc.add_table(rows=len(biz_rules)+1, cols=3)
    rules_tbl.rows[0].cells[0].paragraphs[0].text = "Rule ID"
    rules_tbl.rows[0].cells[1].paragraphs[0].text = "Rule Name"
    rules_tbl.rows[0].cells[2].paragraphs[0].text = "Description & Enforcement Mechanism"
    fill_table(rules_tbl, biz_rules)
    stylize(rules_tbl,[0.8,1.6,4.1])
    doc.add_paragraph()

    # 5. Design Analysis
    head(doc,"5. Design Analysis — Issues Identified and Resolutions")
    issues_tbl = doc.add_table(rows=5, cols=4)
    issues_tbl.rows[0].cells[0].paragraphs[0].text = "Issue ID"
    issues_tbl.rows[0].cells[1].paragraphs[0].text = "Design Issue Identified"
    issues_tbl.rows[0].cells[2].paragraphs[0].text = "Risk / Impact"
    issues_tbl.rows[0].cells[3].paragraphs[0].text = "Resolution Applied"
    issues = [
        ("DI-01","failure_category and failure_reason are denormalized into transactions table (also exist in error_codes)",
         "Data redundancy; inconsistency risk if error_codes table is updated without cascading to transactions",
         "ACCEPTED as a deliberate design choice for Streamlit read performance — avoids a JOIN on every chart render. Source-of-truth is error_codes; transactions.failure_category is a computed copy set at load time."),
        ("DI-02","No explicit bridge table between terminals and error_codes (many-to-many via transactions)",
         "Cannot directly query 'which error codes occur at which terminals' without going through transactions",
         "RESOLVED by the analysis view v_transaction_analysis which provides the terminal↔error_code cross-section. A bridge table would add complexity without analytical value for this scale."),
        ("DI-03","maintenance_logs.component_replaced is nullable TEXT — no reference table for component types",
         "Free-text component names may cause inconsistency in group-by queries (e.g., 'Cash Cassette' vs 'cash cassette')",
         "RESOLVED in generate_data.py by using a fixed controlled vocabulary list for component_replaced values. A formal FK to a components master was deferred to Post-MVP scope."),
        ("DI-04","transactions.issuing_bank is a TEXT column with no FK to a banks master table",
         "Inconsistent bank name strings could corrupt On-Us/Off-Us classification logic",
         "RESOLVED by constraining issuing_bank values to the same 5-bank list as atm_terminals.bank_owner in generate_data.py with a CHECK-equivalent applied at generator level."),
    ]
    fill_table(issues_tbl, issues)
    stylize(issues_tbl,[0.6,2.1,1.6,2.2])
    doc.add_paragraph()

    # 6. Table Column Inventory
    head(doc,"6. Table and Column Quick Reference")
    col_ref = doc.add_table(rows=36, cols=5)
    col_ref.rows[0].cells[0].paragraphs[0].text = "Table"
    col_ref.rows[0].cells[1].paragraphs[0].text = "Column"
    col_ref.rows[0].cells[2].paragraphs[0].text = "Data Type"
    col_ref.rows[0].cells[3].paragraphs[0].text = "Key"
    col_ref.rows[0].cells[4].paragraphs[0].text = "OQ Answered"
    col_rows = [
        # atm_terminals
        ("atm_terminals","terminal_id","TEXT","PK","OQ-03,09,10"),
        ("atm_terminals","bank_owner","TEXT","—","OQ-07,09"),
        ("atm_terminals","vendor","TEXT","—","OQ-09"),
        ("atm_terminals","terminal_model","TEXT","—","OQ-09"),
        ("atm_terminals","install_date","DATE","—","OQ-10"),
        ("atm_terminals","location_name","TEXT","—","OQ-03"),
        ("atm_terminals","city","TEXT","—","OQ-03"),
        ("atm_terminals","state","TEXT","—","—"),
        ("atm_terminals","zone","TEXT","—","OQ-03"),
        ("atm_terminals","status","TEXT","—","OQ-03"),
        ("atm_terminals","latitude","REAL","—","—"),
        ("atm_terminals","longitude","REAL","—","—"),
        # error_codes
        ("error_codes","error_code","TEXT","PK","OQ-02"),
        ("error_codes","iso_name","TEXT","—","OQ-02"),
        ("error_codes","description","TEXT","—","OQ-02"),
        ("error_codes","failure_category","TEXT","—","OQ-02,04"),
        ("error_codes","severity","TEXT","—","OQ-02"),
        ("error_codes","is_customer_fault","INTEGER","—","OQ-05,08"),
        # transactions
        ("transactions","transaction_id","TEXT","PK","OQ-01..10"),
        ("transactions","terminal_id","TEXT","FK","OQ-03,10"),
        ("transactions","timestamp","DATETIME","—","OQ-01,04"),
        ("transactions","transaction_type","TEXT","—","OQ-05"),
        ("transactions","amount","REAL","—","OQ-08"),
        ("transactions","status","TEXT","—","OQ-01..10"),
        ("transactions","error_code","TEXT","FK","OQ-02"),
        ("transactions","response_time_ms","INTEGER","—","OQ-06"),
        ("transactions","card_type","TEXT","—","OQ-07"),
        ("transactions","issuing_bank","TEXT","—","OQ-07"),
        ("transactions","failure_category","TEXT","—","OQ-02,04,08"),
        ("transactions","failure_reason","TEXT","—","OQ-02"),
        # maintenance_logs
        ("maintenance_logs","maintenance_id","TEXT","PK","OQ-10"),
        ("maintenance_logs","terminal_id","TEXT","FK","OQ-10"),
        ("maintenance_logs","event_datetime","DATETIME","—","OQ-10"),
        ("maintenance_logs","event_type","TEXT","—","OQ-10"),
        ("maintenance_logs","cost_inr","REAL","—","OQ-09"),
    ]
    fill_table(col_ref, col_rows)
    stylize(col_ref,[1.4,1.6,1.1,0.6,1.3])
    doc.add_paragraph()

    # Viva Questions
    head(doc,"7. ER Design — Viva Preparation Notes")
    callout(doc,
        "Three Viva Questions: "
        "Q1: Why is the failure_category column duplicated in both error_codes and transactions? "
        "A: Intentional denormalization for Streamlit read performance — avoids a JOIN on every chart render. "
        "Source of truth is error_codes; transactions.failure_category is computed at load time. "
        "Q2: What would happen if PRAGMA foreign_keys = ON was not set in SQLite? "
        "A: SQLite ignores FK constraints by default; orphan transaction records with invalid terminal_id would silently insert, corrupting all terminal-level aggregations. "
        "Q3: Why use a 1:many relationship between error_codes and transactions rather than storing the error description directly in the transactions table? "
        "A: Normalization avoids update anomalies — if the description of error code 91 is corrected, it only needs updating in error_codes, not in 50,000 transaction rows.",
        title="VIVA NOTES")

    doc.save(out)
    print(f"ER Diagram doc saved → {out}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_er_doc("docs/09_ER_Diagram.docx")
