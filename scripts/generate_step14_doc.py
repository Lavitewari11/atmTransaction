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

def build_step14_doc(output_path):
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
    r_sub = p_sub.add_run("Step 14 — Master Build Plan & Work Breakdown Structure  v1.0\n")
    r_sub.font.size = Pt(13)
    r_sub.font.name = "Calibri"
    r_sub.font.color.rgb = RGBColor(74, 85, 104)
    
    meta = doc.add_table(rows=5, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    mdata = [
        ("Project & Code", "p096 | ATM Transaction Failure Analysis"),
        ("Student / Candidate", "Lavi Tewari (Roll: 2500320140085), MCA 2nd Yr DA, ABES"),
        ("Project Mentor", "Mr. Dinesh Sereno"),
        ("Implementation Scope", "BUILD-01 through BUILD-06 Tasks, DoD, and Verification"),
        ("Source Specification", "BUILD_PLAN.md baseline")
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
    
    # Section 1: Executive Summary
    h1 = doc.add_heading("1. Implementation Executive Summary", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    doc.add_paragraph(
        "The Build Plan outlines the staged engineering sequence to translate all preceding architecture, "
        "database, and UI specifications into production code. It enforces an automated test-driven workflow "
        "where database integrity, analytical formulas, and interface components are rigorously verified with pytest."
    )
    doc.add_paragraph()
    
    # Section 2: Build Task Breakdown
    h1 = doc.add_heading("2. Build Task Work Breakdown Structure (WBS)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    b_table = doc.add_table(rows=7, cols=5)
    b_headers = ["Task ID", "Implementation Goal", "Target Files", "Linked Stories", "Verification Test"]
    for i, h in enumerate(b_headers):
        b_table.rows[0].cells[i].paragraphs[0].text = h
    b_data = [
        ("BUILD-01", "SQLite Data Ingestion & Caching Layer", "src/db.py", "US-01 (FR-01, FR-02)", "pytest tests/test_data.py"),
        ("BUILD-02", "Taxonomy Mapping & Data Quality Telemetry", "src/transform.py", "US-02, US-03 (FR-02, FR-03)", "0 unhandled NULL categories"),
        ("BUILD-03", "Statistical Engine & Anomaly Outlier Detection", "src/analysis.py", "US-04, US-10, US-11 (FR-04..13)", "pytest tests/test_analysis.py"),
        ("BUILD-04", "Centralized Plotly Visual Chart Factory", "src/charts.py", "US-04, 07, 08, 10, 13 (FR-14)", "All figures render valid objects"),
        ("BUILD-05", "Streamlit Entrypoint & Sidebar Filter State", "app.py, src/ui.py", "US-01, US-05, US-15 (FR-15, 16)", "Streamlit startup zero errors"),
        ("BUILD-06", "Multipage Diagnostic Dashboards (6 Pages)", "pages/01..06.py", "US-04 through US-14 (FR-04..14)", "Full user journey passes")
    ]
    for r_idx, r_data in enumerate(b_data, start=1):
        for c_idx, text in enumerate(r_data):
            b_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(b_table, [0.8, 1.8, 1.3, 1.2, 1.4])
    doc.add_paragraph()
    
    # Section 3: Definition of Done
    h1 = doc.add_heading("3. Definition of Done (DoD)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    doc.add_paragraph("1. 100% of Must-Have and Should-Have Business Requirements (BR-01..BR-10) functional.")
    doc.add_paragraph("2. All 11 pytest unit tests pass with zero assertion failures.")
    doc.add_paragraph("3. Streamlit application executes locally without traceback exceptions.")
    doc.add_paragraph("4. SQLite database file data/app.db verified with complete foreign key integrity.")
    doc.add_paragraph("5. Pinned requirements.txt and documentation package complete for academic evaluation.")
    doc.add_paragraph()
    
    # Section 4: Viva Questions
    h1 = doc.add_heading("4. Build Plan — Viva Preparation Notes", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    add_callout(doc,
        "TOP 3 BUILD PLAN VIVA DEFENSE QUESTIONS:\n"
        "1. Q: Why sequence data loaders before analytical functions in the WBS?\n"
        "   A: Strict dependency order prevents circular imports and guarantees analytical functions test against valid DataFrame fixtures.\n"
        "2. Q: How does BUILD-02 isolate intentional data quality flaws?\n"
        "   A: The transform pipeline quenches duplicate transaction IDs and bounds extreme amounts (> ₹50,000) while surfacing them in a dedicated telemetry badge.\n"
        "3. Q: How does Streamlit multipage routing maintain active filter state?\n"
        "   A: Sliced DataFrames are stored in st.session_state['filtered_df'], allowing pages 01 to 06 to access pre-filtered data instantaneously.",
        title="VIVA HIGHLIGHTS")
        
    doc.save(output_path)
    print(f"Step 14 DOCX created at: {output_path}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_step14_doc("docs/14_BUILD_PLAN.docx")
