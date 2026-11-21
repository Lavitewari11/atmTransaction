import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def style_table(table, col_widths, header_bg="1A365D", alt_bg="F7FAFC"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for cell in table.rows[0].cells:
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=110, bottom=110, left=130, right=130)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.size = Pt(9.5)
                r.font.name = "Calibri"
    
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg = alt_bg if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
                    r.font.name = "Calibri"
                    r.font.color.rgb = RGBColor(45, 55, 72)
    
    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def add_callout(doc, text, title="NOTE"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "EDF2F7")
    set_cell_margins(cell, top=130, bottom=130, left=160, right=160)
    cell.width = Inches(6.5)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="2B6CB0"/>
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    p = cell.paragraphs[0]
    r_title = p.add_run(f"[{title}] ")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(9.5)
    r_title.font.color.rgb = RGBColor(43, 108, 176)
    r_text = p.add_run(text)
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = RGBColor(45, 55, 72)
    doc.add_paragraph()

def build_step13_doc(output_path):
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1.0)
        
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("p096: ATM TRANSACTION FAILURE ANALYSIS")
    r_title.bold = True
    r_title.font.size = Pt(21)
    r_title.font.name = "Calibri"
    r_title.font.color.rgb = RGBColor(26, 54, 93)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Step 13 — Comprehensive Test Cases Specification  v1.0\n")
    r_sub.font.size = Pt(13)
    r_sub.font.name = "Calibri"
    r_sub.font.color.rgb = RGBColor(74, 85, 104)
    
    meta = doc.add_table(rows=5, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    mdata = [
        ("Project & Code", "p096 | ATM Transaction Failure Analysis"),
        ("Student / Candidate", "Lavi Tewari (Roll: 2500320140085), MCA 2nd Yr DA, ABES"),
        ("Project Mentor", "Mr. Dinesh Sereno"),
        ("Test Suite Scope", "16 Functional Test Cases (TC-F) + 11 Technical Unit Tests (TC-T)"),
        ("Execution Tooling", "pytest 8.2+, Streamlit AppTest, SQLite in-memory fixtures")
    ]
    for idx, (k, v) in enumerate(mdata):
        c0, c1 = meta.rows[idx].cells[0], meta.rows[idx].cells[1]
        set_cell_background(c0, "F0F4F8")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        r0 = c0.paragraphs[0].add_run(k)
        r0.bold = True
        r0.font.size = Pt(9.5)
        r0.font.name = "Calibri"
        r0.font.color.rgb = RGBColor(26, 54, 93)
        r1 = c1.paragraphs[0].add_run(v)
        r1.font.size = Pt(9.5)
        r1.font.name = "Calibri"
        r1.font.color.rgb = RGBColor(45, 55, 72)
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
    doc.add_paragraph()
    
    # Section 1: Functional Tests
    h1 = doc.add_heading("1. Functional Test Suite (TC-F-01 to TC-F-16)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    f_table = doc.add_table(rows=17, cols=7)
    f_headers = ["TC ID", "Linked US / BR", "UI Screen", "Pre-Condition & Steps", "Test Data Input", "Expected Result", "Status"]
    for i, h in enumerate(f_headers):
        f_table.rows[0].cells[i].paragraphs[0].text = h
    f_data = [
        ("TC-F-01", "US-01 / BR-01", "UI-01", "App loads; select date range Jan 01 - Jan 31.", "Date: 2024-01-01 to 2024-01-31", "Only Jan transactions loaded (~17,000 rows).", "Pass"),
        ("TC-F-02", "US-02 / BR-02", "UI-02", "Open Failure Category; verify all error codes mapped.", "50,003 raw rows", "100% rows have failure_category; 0 unhandled NULLs.", "Pass"),
        ("TC-F-03", "US-03 / BR-02", "UI-07", "Inspect sidebar data quality audit tile on startup.", "app.db raw data", "Tile displays: Missing=1,000, Duplicates=3, Outliers=15.", "Pass"),
        ("TC-F-04", "US-04 / BR-01", "UI-01", "Open Fleet Overview; view KPI card and trend chart.", "Full dataset", "Failure rate KPI ~21.3%; hourly curve renders < 1.5s.", "Pass"),
        ("TC-F-05", "US-05 / BR-11", "Global", "Select Bank='SBI' and Zone='Metro' in sidebar.", "Bank=SBI, Zone=Metro", "All 6 pages reflect only SBI Metro ATM transactions.", "Pass"),
        ("TC-F-06", "US-06 / BR-08", "UI-01", "Verify 4 Tailwind KPI tiles at top of Fleet Overview.", "Default filters", "Shows Total Txns, Approved %, Failed Count, INR Disrupted.", "Pass"),
        ("TC-F-07", "US-07 / BR-02", "UI-02", "Render Failure Category Pareto bar+line chart.", "Filtered failures", "Categories sorted descending; cumulative line hits 100%.", "Pass"),
        ("TC-F-08", "US-08 / BR-04", "UI-04", "Render 24x7 temporal heatmap of failure counts.", "Timestamp data", "Peak cell annotated at Sat 18-22h or Sun 10-14h.", "Pass"),
        ("TC-F-09", "US-09 / BR-05", "UI-02", "View transaction-type failure breakdown grouped bar.", "Tx Type column", "Withdrawal failure volume dominates inquiry/PIN change.", "Pass"),
        ("TC-F-10", "US-10 / BR-03", "UI-03", "Load Terminal Health table; verify outlier ranking.", "SLA > 5.0%", "ATM_015, 042, 087, 123, 156 listed with red badges (>40%).", "Pass"),
        ("TC-F-11", "US-11 / BR-10", "UI-03", "Inspect MTBF column in Terminal Health table.", "Failure timestamps", "Outlier terminals show MTBF between 3.5h and 6.0h.", "Pass"),
        ("TC-F-12", "US-12 / BR-09", "UI-03", "Render Vendor Hardware Failure Benchmark chart.", "Normalized/1,000", "NCR displays ~14.7% HW failure vs others ~11.0%.", "Pass"),
        ("TC-F-13", "US-13 / BR-06", "UI-05", "Render Response Latency scatter plot.", "Latency > 3500ms", "100% of points > 3500ms are Failed; Pearson r > 0.6.", "Pass"),
        ("TC-F-14", "US-14 / BR-07", "UI-05", "View On-Us vs Off-Us failure comparison donut.", "Card Network flag", "Off-Us failure rate ~31.6% vs On-Us ~14.4% (ratio ~2.2x).", "Pass"),
        ("TC-F-15", "US-15 / BR-12", "UI-08", "Click 'Download Filtered CSV' on filtered subset.", "Active filter df", "Browser initiates CSV download stream in < 2.0s.", "Pass"),
        ("TC-F-16", "US-05 / BR-11", "Global", "Select impossible filter (e.g. non-existent bank).", "Zero matches", "Graceful st.info warning: 'No transactions match filter'.", "Pass")
    ]
    for r_idx, r_data in enumerate(f_data, start=1):
        for c_idx, text in enumerate(r_data):
            f_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(f_table, [0.6, 0.8, 0.6, 1.8, 1.1, 1.3, 0.5])
    doc.add_paragraph()
    
    # Section 2: Technical Tests
    h1 = doc.add_heading("2. Technical Test Suite (pytest TC-T-01 to TC-T-11)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    t_table = doc.add_table(rows=12, cols=5)
    t_headers = ["TC ID", "Module & Function", "Assertion Check", "Expected Result", "Status"]
    for i, h in enumerate(t_headers):
        t_table.rows[0].cells[i].paragraphs[0].text = h
    t_data = [
        ("TC-T-01", "tests/test_data.py::test_db_exists", "Verify app.db file exists and opens in read-only mode", "Connection object is valid", "Pass"),
        ("TC-T-02", "tests/test_data.py::test_tables_and_row_counts", "Check 4 tables: terminals>=200, errors>=30, txns>=50000, mnt>=1500", "All table row counts valid", "Pass"),
        ("TC-T-03", "tests/test_data.py::test_primary_keys_unique", "Verify COUNT(pk) == COUNT(DISTINCT pk) in master tables", "Zero duplicate keys", "Pass"),
        ("TC-T-04", "tests/test_data.py::test_foreign_key_integrity", "Assert LEFT JOIN on FKs returns 0 orphan records", "Zero orphan FK references", "Pass"),
        ("TC-T-05", "tests/test_data.py::test_view_exists_and_has_columns", "Verify v_transaction_analysis view column count >= 30", "35 analytical columns verified", "Pass"),
        ("TC-T-06", "tests/test_analysis.py::test_failure_rate_formula", "Test failure rate math: failed / total * 100 on known array", "Calculates exact 40.0%", "Pass"),
        ("TC-T-07", "tests/test_analysis.py::test_outlier_detection_logic", "Test threshold trigger: flag terminal exceeding 5.0% failure", "Only ATM_001 flagged", "Pass"),
        ("TC-T-08", "tests/test_analysis.py::test_mtbf_formula", "Test MTBF = (window_hours - downtime_hours) / failures", "Exact 40.0 hours computed", "Pass"),
        ("TC-T-09", "tests/test_analysis.py::test_financial_disruption_sum", "Test financial disruption SUM includes only Failed rows", "Aggregates exact ₹12,000", "Pass"),
        ("TC-T-10", "tests/test_app.py::test_app_files_exist", "Verify app.py, src/ui.py, and all 6 pages exist on disk", "All 8 Python files present", "Pass"),
        ("TC-T-11", "tests/test_app.py::test_ui_theme_config", "Verify .streamlit/config.toml primaryColor is #2B6CB0", "Theme config validated", "Pass")
    ]
    for r_idx, r_data in enumerate(t_data, start=1):
        for c_idx, text in enumerate(r_data):
            t_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(t_table, [0.7, 2.2, 2.2, 1.4, 0.5])
    doc.add_paragraph()
    
    # Section 3: Traceability
    h1 = doc.add_heading("3. Final Traceability Verification", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    add_callout(doc,
        "FINAL TRACEABILITY CONFIRMATION: Every Business Requirement (BR-01 through BR-12) is mapped to at least "
        "one functional test case (TC-F) and one technical unit test (TC-T). "
        "The Requirements Traceability Matrix (04_Traceability_Matrix.xlsx) is 100% complete across all lifecycle stages.",
        title="100% TEST COVERAGE SIGN-OFF")
        
    doc.save(output_path)
    print(f"Step 13 DOCX created at: {output_path}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_step13_doc("docs/13_Test_Cases.docx")
