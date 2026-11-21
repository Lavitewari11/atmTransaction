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

def build_step11_doc(output_path):
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
    r_sub = p_sub.add_run("Step 11 — Agile User Stories & Acceptance Criteria  v1.0\n")
    r_sub.font.size = Pt(13)
    r_sub.font.name = "Calibri"
    r_sub.font.color.rgb = RGBColor(74, 85, 104)
    
    meta = doc.add_table(rows=5, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    mdata = [
        ("Project & Code", "p096 | ATM Transaction Failure Analysis"),
        ("Student / Candidate", "Lavi Tewari (Roll: 2500320140085), MCA 2nd Yr DA, ABES"),
        ("Project Mentor", "Mr. Dinesh Sereno"),
        ("Agile Framework", "Scrum / 1-Week Sprint (Story Points: Fibonacci 1, 2, 3, 5, 8)"),
        ("Document Baseline", "v1.0 (Mapped to BR-01..BR-12 and FR-01..FR-16)")
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
    
    # Section 1: Epic Summary
    h1 = doc.add_heading("1. Epic Decomposition & Story Point Summary", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    epic_table = doc.add_table(rows=6, cols=5)
    epic_headers = ["Epic ID", "Epic Name", "Enclosed User Stories", "Total Points", "Strategic Objective"]
    for i, h in enumerate(epic_headers):
        epic_table.rows[0].cells[i].paragraphs[0].text = h
    epic_data = [
        ("EP-01", "Data Foundation & Quality", "US-01, US-02, US-03", "8 pts", "Reliable SQLite ingestion, cached schema validation, and data hygiene audits."),
        ("EP-02", "Fleet Overview Dashboard", "US-04, US-05, US-06", "10 pts", "Executive KPI summary rows, interactive trend curves, and global sidebar filter state."),
        ("EP-03", "Root-Cause Failure Analysis", "US-07, US-08, US-09", "13 pts", "Categorical Pareto analysis, 24x7 congestion heatmap, and transaction type breakdown."),
        ("EP-04", "Terminal Health & Outliers", "US-10, US-11, US-12", "13 pts", "Chronic terminal SLA breach ranking (>5%), MTBF calculations, and vendor benchmarking."),
        ("EP-05", "Network Latency & Export", "US-13, US-14, US-15", "10 pts", "Statistical latency correlation (Pearson r), On-Us/Off-Us Z-test, and CSV export.")
    ]
    for r_idx, r_data in enumerate(epic_data, start=1):
        for c_idx, text in enumerate(r_data):
            epic_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(epic_table, [0.8, 1.8, 1.4, 0.8, 2.2])
    doc.add_paragraph()
    
    # Section 2: Master User Story Table
    h1 = doc.add_heading("2. Master User Stories Table (US-01 to US-15)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    us_table = doc.add_table(rows=16, cols=8)
    us_headers = ["US ID", "Epic", "Persona", "User Story Statement", "Feature ID", "BR ID", "Pri", "Pts"]
    for i, h in enumerate(us_headers):
        us_table.rows[0].cells[i].paragraphs[0].text = h
        
    master_stories = [
        ("US-01", "EP-01", "ATM Ops Manager", "As an ATM Ops Manager, I want 50,000 transaction records loaded and cached with date filtering so that I see only relevant operational periods.", "FR-01, 15", "BR-01, 11", "Must", "3"),
        ("US-02", "EP-01", "NOC Analyst", "As a NOC Analyst, I want error codes mapped automatically to the 4-tier banking taxonomy so that I never see unclassified transactions.", "FR-02, 03", "BR-02", "Must", "3"),
        ("US-03", "EP-01", "Field Lead", "As a Field Maintenance Lead, I want a startup data-quality audit counting missing and duplicate records so that I trust the analytics.", "FR-02", "BR-02", "Should", "2"),
        ("US-04", "EP-02", "ATM Ops Manager", "As an ATM Ops Manager, I want fleet failure rate displayed as a KPI card and hourly trend line so that I can spot intraday spikes instantly.", "FR-04, 14", "BR-01", "Must", "5"),
        ("US-05", "EP-02", "ATM Ops Manager", "As an ATM Ops Manager, I want multi-select filters for bank, zone, and transaction type so that I can isolate operational subsets.", "FR-15", "BR-11", "Could", "3"),
        ("US-06", "EP-02", "COO", "As a COO, I want an executive summary row showing total transactions, approval rate, and INR at risk so that I see macro business impact.", "FR-04, 11", "BR-01, 08", "Must", "2"),
        ("US-07", "EP-03", "NOC Analyst", "As a NOC Analyst, I want a Pareto chart ranking failure categories by volume so that I target the top categories driving 80% of declines.", "FR-07, 14", "BR-02, 05", "Must", "5"),
        ("US-08", "EP-03", "NOC Analyst", "As a NOC Analyst, I want a 24-hour x 7-day temporal heatmap of failure counts so that I can anticipate network congestion peaks.", "FR-06, 14", "BR-04", "Should", "5"),
        ("US-09", "EP-03", "Dispute Officer", "As a Dispute Officer, I want failure rates segmented across transaction types so that I can prioritize high-friction cash dispense failures.", "FR-07, 14", "BR-05", "Must", "3"),
        ("US-10", "EP-04", "Field Lead", "As a Field Lead, I want a ranked table of outlier terminals exceeding 5% failure SLA so that I can dispatch engineers to chronic nodes.", "FR-05, 14", "BR-03", "Must", "5"),
        ("US-11", "EP-04", "Field Lead", "As a Field Lead, I want MTBF (in hours) computed for each outlier terminal so that I can schedule preventive repairs before complete breakdown.", "FR-13, 14", "BR-10", "Must", "5"),
        ("US-12", "EP-04", "COO", "As a COO, I want hardware failure rates benchmarked per 1,000 transactions across vendors so that I can hold hardware vendors accountable.", "FR-12, 14", "BR-09", "Should", "3"),
        ("US-13", "EP-05", "NOC Analyst", "As a NOC Analyst, I want a response time scatter plot with Pearson correlation so that I can statistically prove network latency causes timeouts.", "FR-08, 09, 14", "BR-06", "Must", "5"),
        ("US-14", "EP-05", "ATM Ops Manager", "As an ATM Ops Manager, I want On-Us vs Off-Us failure rates compared with a Z-test so that I can prove whether external switches are bottlenecked.", "FR-10, 14", "BR-07", "Should", "3"),
        ("US-15", "EP-05", "Field Lead", "As a Field Lead, I want to download filtered transaction subsets as CSV/Excel so that field technicians have offline inspection sheets on site.", "FR-16", "BR-12", "Could", "2")
    ]
    for r_idx, r_data in enumerate(master_stories, start=1):
        for c_idx, text in enumerate(r_data):
            us_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(us_table, [0.6, 0.6, 1.0, 2.4, 0.7, 0.6, 0.5, 0.4])
    doc.add_paragraph()
    
    # Section 3: Detailed Acceptance Criteria
    h1 = doc.add_heading("3. Detailed User Stories & Given-When-Then Acceptance Criteria", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    detailed_stories = [
        ("US-01", "Cached Ingestion & Date Slicing", "ATM Operations Manager", "Must", "3 pts", "FR-01, FR-15", "BR-01, BR-11", "OQ-01",
         "Given the SQLite analytical database is connected at runtime,\n"
         "When the user selects a custom date range in the sidebar,\n"
         "Then the application loads cached records within 2.0s and slices DataFrames without full table re-querying."),
        ("US-02", "4-Tier Banking Taxonomy Mapping", "NOC & Switch Analyst", "Must", "3 pts", "FR-02, FR-03", "BR-02", "OQ-02",
         "Given transaction records contain standard ISO 8583 response codes,\n"
         "When the transformation pipeline executes,\n"
         "Then 100% of rows resolve to Hardware, Network-Switch, Customer-Account, or Host-Timeout with 0 unhandled NULLs."),
        ("US-03", "Data Quality Audit Telemetry", "Field Maintenance Lead", "Should", "2 pts", "FR-02", "BR-02", "OQ-02",
         "Given the synthetic dataset has embedded missing values, duplicates, and outliers,\n"
         "When the diagnostic module initializes,\n"
         "Then a dedicated sidebar audit badge displays exact counts of quarantined anomaly records."),
        ("US-04", "Fleet Failure Trend Curve & KPI", "ATM Operations Manager", "Must", "5 pts", "FR-04, FR-14", "BR-01", "OQ-01",
         "Given the user opens the Fleet Overview page,\n"
         "When the date filter is modified,\n"
         "Then the Fleet Failure Rate KPI card and the hourly Plotly trend curve refresh simultaneously in under 1.5 seconds."),
        ("US-05", "Multi-Dimensional Filter Engine", "ATM Operations Manager", "Could", "3 pts", "FR-15", "BR-11", "OQ-01, OQ-03",
         "Given multiple criteria (Bank, City Zone, Transaction Type) are selected in the sidebar,\n"
         "When any dropdown selection changes,\n"
         "Then session state propagates cleanly across all 6 multipage analytical views without state reset."),
        ("US-06", "Executive Financial & Operational Summary", "Chief Operations Officer (COO)", "Must", "2 pts", "FR-04, FR-11", "BR-01, BR-08", "OQ-01, OQ-08",
         "Given executive management accesses the platform,\n"
         "When reviewing high-level portfolio health,\n"
         "Then 4 Tailwind KPI tiles immediately present Total Volume, Approval %, Failed Count, and Gross INR Disruption."),
        ("US-07", "Pareto 80/20 Failure Root-Cause Analysis", "NOC & Switch Analyst", "Must", "5 pts", "FR-07, FR-14", "BR-02, BR-05", "OQ-02",
         "Given categorical failure data is aggregated,\n"
         "When opening the Failure Category screen,\n"
         "Then a dual-axis Plotly Pareto chart renders category bars in descending order overlaid with an 80% cumulative curve."),
        ("US-08", "24x7 Temporal Congestion Heatmap", "NOC & Switch Analyst", "Should", "5 pts", "FR-06, FR-14", "BR-04", "OQ-04",
         "Given timestamp telemetry with hour-of-day and day-of-week dimensions,\n"
         "When the Temporal Patterns page renders,\n"
         "Then a 24x7 matrix clearly visualizes the Saturday 6-10 PM and Sunday 10 AM-2 PM failure congestion peaks."),
        ("US-09", "Transaction-Type Failure Segmentation", "Customer Dispute Officer", "Must", "3 pts", "FR-07, FR-14", "BR-05", "OQ-05",
         "Given disputes require separating cash-trap claims from balance inquiries,\n"
         "When viewing the Transaction Type breakdown,\n"
         "Then a grouped bar chart verifies Cash-Withdrawal volume vs PIN-Change and Balance Inquiry failure rates."),
        ("US-10", "Chronic Outlier Terminal Ranking (>5% SLA)", "Field Maintenance Lead", "Must", "5 pts", "FR-05, FR-14", "BR-03", "OQ-03",
         "Given terminal failure rates are computed over a 7-day rolling window,\n"
         "When ranking terminals descending by failure rate,\n"
         "Then chronic machines (ATM_015, 042, 087, 123, 156) appear prominently with red warning badges."),
        ("US-11", "MTBF Predictive Maintenance Metric", "Field Maintenance Lead", "Must", "5 pts", "FR-13, FR-14", "BR-10", "OQ-10",
         "Given consecutive failure timestamps and maintenance event logs,\n"
         "When calculating operating hours between failures,\n"
         "Then MTBF is displayed per terminal and validated against hand-calculated test fixtures within +/- 5%."),
        ("US-12", "Hardware Manufacturer Reliability Benchmarking", "Chief Operations Officer (COO)", "Should", "3 pts", "FR-12, FR-14", "BR-09", "OQ-09",
         "Given terminal vendor telemetry across NCR, Diebold, Hitachi, GRG, and Nautilus,\n"
         "When viewing the Vendor Benchmark chart,\n"
         "Then failure rates are normalized per 1,000 transactions, exposing vendor-specific hardware vulnerability."),
        ("US-13", "Network Latency vs Timeout Correlation", "NOC & Switch Analyst", "Must", "5 pts", "FR-08, FR-09, FR-14", "BR-06", "OQ-06",
         "Given response latency in milliseconds is recorded,\n"
         "When plotting latency against transaction outcome,\n"
         "Then a Plotly scatter plot and Pearson correlation coefficient prove latency > 3500ms triggers near-certain timeout."),
        ("US-14", "On-Us vs Off-Us Switch Differential & Z-Test", "ATM Operations Manager", "Should", "3 pts", "FR-10, FR-14", "BR-07", "OQ-07",
         "Given card interchange networks experience gateway drops,\n"
         "When comparing On-Us vs Off-Us failure proportions,\n"
         "Then a donut chart displays volume split and a two-proportion Z-test proves statistical significance (p < 0.001)."),
        ("US-15", "Export Filtered Audit Subsets to CSV", "Field Maintenance Lead", "Could", "2 pts", "FR-16", "BR-12", "OQ-03, OQ-08",
         "Given field engineers require offline terminal error sheets,\n"
         "When clicking the 'Download Filtered CSV' button,\n"
         "Then a clean CSV extract containing all visible columns downloads in under 2.0 seconds.")
    ]
    
    for s_id, s_title, s_pers, s_prio, s_pts, s_feat, s_br, s_oq, ac in detailed_stories:
        head = doc.add_heading(f"{s_id}: {s_title}", level=2)
        head.runs[0].font.color.rgb = RGBColor(26, 54, 93)
        
        card = doc.add_table(rows=5, cols=2)
        cdata = [
            ("User Persona & Role", f"{s_pers} (Priority: {s_prio} | Estimate: {s_pts})"),
            ("System Traceability", f"Feature: {s_feat} | Requirement: {s_br} | Operational Question: {s_oq}"),
            ("User Story Statement", f"\"As a {s_pers}, I want to {s_title.lower()} so that operational failure risks are mitigated.\""),
            ("Acceptance Criteria (G/W/T)", ac),
            ("Definition of Done (DoD)", "Code implemented in src/, unit test passing in pytest, rendered in Streamlit UI with zero console errors.")
        ]
        for idx, (label, val) in enumerate(cdata):
            c0, c1 = card.rows[idx].cells[0], card.rows[idx].cells[1]
            set_cell_background(c0, "EBF4FF")
            set_cell_background(c1, "FFFFFF")
            set_cell_margins(c0, 60, 60, 100, 100)
            set_cell_margins(c1, 60, 60, 100, 100)
            r0 = c0.paragraphs[0].add_run(label)
            r0.bold = True
            r0.font.size = Pt(9)
            r0.font.name = "Calibri"
            r0.font.color.rgb = RGBColor(26, 54, 93)
            r1 = c1.paragraphs[0].add_run(val)
            r1.font.size = Pt(9)
            r1.font.name = "Calibri"
            r1.font.color.rgb = RGBColor(45, 55, 72)
            c0.width = Inches(1.8)
            c1.width = Inches(4.7)
        style_table(card, [1.8, 4.7])
        doc.add_paragraph()
        
    # Section 4: Sprint Plan
    h1 = doc.add_heading("4. 5-Day Sprint Execution Plan", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    sprint_table = doc.add_table(rows=6, cols=5)
    sp_headers = ["Sprint Day", "Planned Stories", "Operational Focus Area", "Points", "Deliverable Verification"]
    for i, h in enumerate(sp_headers):
        sprint_table.rows[0].cells[i].paragraphs[0].text = h
    sp_data = [
        ("Day 1 (Mon)", "US-01, US-02, US-03", "Data Layer & Quality Pipeline", "8 pts", "SQLite connector, taxonomy mapper, and anomaly badge validated."),
        ("Day 2 (Tue)", "US-04, US-06", "Fleet Overview & KPI Summary", "7 pts", "Tailwind KPI row + hourly failure trend curve rendering < 1.5s."),
        ("Day 3 (Wed)", "US-07, US-09", "Failure Taxonomy & Segmentation", "8 pts", "Pareto 80/20 chart + transaction type grouped breakdown operational."),
        ("Day 4 (Thu)", "US-10, US-11, US-05", "Outlier Ranking & MTBF Module", "13 pts", "Chronic outlier table (>5% SLA) + MTBF calculations + global filter sync."),
        ("Day 5 (Fri)", "US-08, US-12, US-13, US-14, US-15", "Temporal, Latency & Data Export", "18 pts", "24x7 heatmap + latency scatter + Z-test donut + CSV download stream.")
    ]
    for r_idx, r_data in enumerate(sp_data, start=1):
        for c_idx, text in enumerate(r_data):
            sprint_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(sprint_table, [1.0, 1.4, 1.8, 0.7, 1.6])
    doc.add_paragraph()
    
    # Section 5: RTM Traceability Update
    h1 = doc.add_heading("5. Traceability Mapping: User Stories to Business Requirements", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    add_callout(doc,
        "RTM UPDATE CONFIRMATION: The 15 user stories (US-01 through US-15) mapped directly across BR-01 through BR-12. "
        "Every Must-Have requirement has at least one dedicated user story with verifiable Given-When-Then acceptance criteria. "
        "Total sprint velocity: 54 story points across a structured 5-day academic build cycle.",
        title="TRACEABILITY VERIFIED")
        
    doc.save(output_path)
    print(f"Step 11 DOCX created at: {output_path}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_step11_doc("docs/11_User_Stories.docx")
