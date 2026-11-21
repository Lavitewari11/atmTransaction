import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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
    # Format header row
    header_tr = table.rows[0]._tr.get_or_add_trPr()
    header_tr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for i, cell in enumerate(table.rows[0].cells):
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.size = Pt(9.5)
                r.font.name = "Calibri"
    
    # Format data rows
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg = alt_bg if r_idx % 2 == 1 else "FFFFFF"
        for i, cell in enumerate(row.cells):
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
                    r.font.name = "Calibri"
                    r.font.color.rgb = RGBColor(45, 55, 72)
    
    # Set widths
    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def add_callout(doc, text, title="NOTE"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "EDF2F7")
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    cell.width = Inches(6.5)
    
    # Left border highlight
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

def build_step1_document(output_path):
    doc = Document()
    
    # Set page margins (1 inch around)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
    
    # Title Block
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("p096: ATM TRANSACTION FAILURE ANALYSIS")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = RGBColor(26, 54, 93) # Deep Navy
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Step 1 — Project Understanding & Operational Baselines\n")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(74, 85, 104)
    
    # Metadata Box Table
    meta_table = doc.add_table(rows=7, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Project Code & Title", "p096 | ATM Transaction Failure Analysis"),
        ("Domain", "Banking & Retail Payment Systems / Data Analytics"),
        ("Student / Candidate", "Lavi Tewari (Roll No: 2500320140085)"),
        ("Academic Program", "MCA 2nd Year (Data Analytics Specialization), ABES Engineering College"),
        ("Project Mentor", "Mr. Dinesh Sereno"),
        ("Target Platform / Stack", "Python, Streamlit, SQLite, Plotly, pytest, Streamlit Community Cloud"),
        ("Document Status & Date", "Version 1.0 (Final Approved) | November 2026")
    ]
    
    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        set_cell_background(cell_lbl, "F0F4F8")
        set_cell_background(cell_val, "FFFFFF")
        set_cell_margins(cell_lbl, top=80, bottom=80, left=120, right=120)
        set_cell_margins(cell_val, top=80, bottom=80, left=120, right=120)
        
        p0 = cell_lbl.paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(9.5)
        r0.font.name = "Calibri"
        r0.font.color.rgb = RGBColor(26, 54, 93)
        
        p1 = cell_val.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.size = Pt(9.5)
        r1.font.name = "Calibri"
        r1.font.color.rgb = RGBColor(45, 55, 72)
        
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.3)
    
    doc.add_paragraph() # Spacing
    
    # SECTION 1: Executive Summary
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Executive Summary")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    p = doc.add_paragraph(
        "Automated Teller Machine (ATM) networks represent critical consumer touchpoints in the retail banking "
        "ecosystem, handling high-volume cash dispensation, balance inquiries, and interbank switch operations. "
        "However, transaction failures resulting from terminal hardware malfunctions, communication switch dropouts, "
        "insufficient cash reserves, and core banking host timeouts inflict substantial operational friction. "
        "This project, designated as p096, develops an interactive, data-driven analytical application that ingests "
        "synthetic multi-bank ATM operational logs, categorizes failure signatures using standard banking failure taxonomy, "
        "detects operational anomalies across terminal clusters, and empowers banking support teams to execute timely "
        "remediation, thereby safeguarding service-level agreements (SLAs) and customer trust."
    )
    p.style.font.name = "Calibri"
    p.style.font.size = Pt(10)
    
    # SECTION 2: Problem Statement
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("2. Problem Statement")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    add_callout(
        doc,
        "\"ATM terminal downtime and recurring transaction failures (technical timeouts, cash dispenser jams, "
        "switch link outages, and customer authentication errors) reduce transaction success rates below acceptable "
        "banking service-level agreements (SLAs), erode customer trust, and inflate field operational maintenance costs "
        "due to the absence of centralized, real-time diagnostic visibility.\"",
        title="CORE PROBLEM STATEMENT"
    )
    
    p = doc.add_paragraph(
        "In modern multi-vendor ATM environments, operational and transaction logs remain dispersed across isolated "
        "switches (NPCI / NFS gateways) and proprietary ATM terminal controllers. Support teams struggle with reactive "
        "troubleshooting, where terminal faults are discovered only after repeated customer grievances or prolonged device "
        "blackouts. A unified analytical platform is required to synthesize transaction logs into actionable failure "
        "intelligence, allowing engineering and operations teams to transition from reactive dispute handling to proactive "
        "terminal fleet optimization."
    )
    p.style.font.name = "Calibri"
    
    # SECTION 3: Intended Audience & User Personas
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("3. Intended Audience and User Personas")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    doc.add_paragraph("The application serves primary and secondary operational roles within retail banking IT and fleet management:")
    
    aud_table = doc.add_table(rows=6, cols=4)
    aud_headers = ["Persona / User Group", "Department / Role", "Usage Frequency", "Primary Interaction & Objective"]
    for i, h in enumerate(aud_headers):
        aud_table.rows[0].cells[i].paragraphs[0].text = h
        
    aud_data = [
        ("ATM Operations Manager", "Retail Banking Fleet Ops", "Daily (Morning & Peak Shifts)", "Reviews fleet-wide availability KPIs, identifies SLA breach hotspots, and re-allocates maintenance shifts."),
        ("Lead Field Service Engineer", "Hardware Vendor / MSP", "Daily (Ad-hoc Dispatch)", "Examines localized hardware failure codes (dispenser jams, card reader errors) to route field technicians."),
        ("NOC / Switch Network Analyst", "Central IT Operations", "Continuous / Hourly", "Monitors switch timeout rates, inter-bank gateway link health, and isolates core banking host disconnects."),
        ("Customer Dispute / QA Officer", "Banking Support Services", "Weekly / On Demand", "Audits chargeback risks, verifies customer-side failure claims (PIN error vs. hardware cash trap), and resolves disputes."),
        ("Branch Banking Head", "Regional Branch Network", "Weekly / Monthly", "Assesses terminal profitability, customer retention impact, and downtime trends across branch-attached ATMs.")
    ]
    
    for row_idx, row_data in enumerate(aud_data, start=1):
        for col_idx, text in enumerate(row_data):
            aud_table.rows[row_idx].cells[col_idx].paragraphs[0].text = text
            
    style_table(aud_table, [1.5, 1.4, 1.3, 2.3])
    doc.add_paragraph()
    
    # SECTION 4: Stakeholder Analysis
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("4. Stakeholder Analysis")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    p = doc.add_paragraph("Detailed stakeholder mapping identifying organizational influence, interest, and critical decisions driven by the analytical dashboard:")
    
    sh_table = doc.add_table(rows=7, cols=6)
    sh_headers = ["SH ID", "Stakeholder", "Organizational Role", "Interest", "Influence", "Decision Made With App"]
    for i, h in enumerate(sh_headers):
        sh_table.rows[0].cells[i].paragraphs[0].text = h
        
    sh_data = [
        ("SH-01", "ATM Operations Manager", "Fleet Operations Lead", "High", "High", "Reallocates cash replenishment cycles, re-routes transaction routing, and enforces vendor SLA penalties."),
        ("SH-02", "Field Maintenance Lead", "Hardware Support & MSP Lead", "High", "Medium", "Prioritizes urgent on-site field visits, orders replacement spare parts (dispensers, sensors), and schedules PM cycles."),
        ("SH-03", "NOC & Switch Analyst", "Switch Network Engineer", "High", "Medium", "Adjusts switch response timeouts, diagnoses telecom gateway packet losses, and escalates switch vendor outages."),
        ("SH-04", "Customer Dispute Officer", "Customer Support & Grievance", "Medium", "Low", "Authorizes automated claim reconciliations, verifies dispense failure veracity, and addresses regulatory consumer complaints."),
        ("SH-05", "Chief Operations Officer (COO)", "Executive Management Sponsor", "Low", "High", "Allocates annual capital budget for fleet modernization, approves ATM vendor contracts, and reviews macro SLA compliance."),
        ("SH-06", "Academic Project Mentor", "Faculty Guide (Mr. Dinesh Sereno)", "High", "High", "Evaluates analytical methodology, software architecture hygiene, data modeling accuracy, and ABES curriculum criteria.")
    ]
    
    for row_idx, row_data in enumerate(sh_data, start=1):
        for col_idx, text in enumerate(row_data):
            sh_table.rows[row_idx].cells[col_idx].paragraphs[0].text = text
            
    style_table(sh_table, [0.7, 1.2, 1.3, 0.7, 0.7, 1.9])
    doc.add_paragraph()
    
    # SECTION 5: Operational Questions (OQs)
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("5. Operational Questions (OQs) & Data Feasibility")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    p = doc.add_paragraph(
        "The system must answer the following 10 core operational questions (OQ-01 through OQ-10). "
        "Each operational question has been verified against the underlying transactional data attributes to guarantee analytical feasibility:"
    )
    
    oq_table = doc.add_table(rows=11, cols=4)
    oq_headers = ["OQ ID", "Operational Question", "Target Stakeholder", "Required Data Attributes"]
    for i, h in enumerate(oq_headers):
        oq_table.rows[0].cells[i].paragraphs[0].text = h
        
    oq_data = [
        ("OQ-01", "What is the overall fleet-wide failure rate, and how does it trend over hourly, daily, and weekly intervals?", "SH-01, SH-05", "timestamp, status, transaction_id"),
        ("OQ-02", "Which failure categories (Hardware, Switch/Network, Customer, Host Timeout) account for the greatest share of decline volume?", "SH-01, SH-03", "failure_category, error_code, status"),
        ("OQ-03", "Which specific ATM terminals consistently exceed the critical 5% failure threshold and qualify as chronic problem nodes?", "SH-01, SH-02", "terminal_id, location, status, error_code"),
        ("OQ-04", "At what specific hours of the day or days of the week do switch-level timeouts and communication drops spike?", "SH-03", "timestamp (hour, day), error_code, failure_category"),
        ("OQ-05", "How does transaction failure incidence vary across different transaction types (Cash Withdrawal, Balance Inquiry, PIN Change)?", "SH-01, SH-04", "transaction_type, status, failure_reason"),
        ("OQ-06", "What is the relationship between system response latency (round-trip time in ms) and ultimate transaction timeout failure?", "SH-03", "response_time_ms, status, error_code"),
        ("OQ-07", "Are transaction declines predominantly concentrated in On-Us (same bank) or Off-Us (interchange network) transactions?", "SH-01, SH-03", "card_type (on-us/off-us), issuing_bank, status"),
        ("OQ-08", "What total financial transaction value (in INR) is disrupted by cash dispenser faults versus customer authorization errors?", "SH-01, SH-04", "amount, failure_category, status"),
        ("OQ-09", "Which ATM hardware manufacturers and terminal models experience disproportionate component failure rates?", "SH-02, SH-05", "terminal_id, terminal_model, failure_category"),
        ("OQ-10", "What is the Mean Time Between Failures (MTBF) and recurrent outage frequency for terminals identified with intermittent connectivity?", "SH-02, SH-03", "terminal_id, timestamp, status, error_code")
    ]
    
    for row_idx, row_data in enumerate(oq_data, start=1):
        for col_idx, text in enumerate(row_data):
            oq_table.rows[row_idx].cells[col_idx].paragraphs[0].text = text
            
    style_table(oq_table, [0.7, 2.6, 1.2, 2.0])
    doc.add_paragraph()
    
    # SECTION 6: Project Objectives (SMART)
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("6. Project Objectives (SMART Criteria)")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    obj_table = doc.add_table(rows=6, cols=3)
    obj_headers = ["Objective ID", "SMART Objective Statement", "SMART Criteria Validation"]
    for i, h in enumerate(obj_headers):
        obj_table.rows[0].cells[i].paragraphs[0].text = h
        
    obj_data = [
        ("OBJ-01", "Design and build an interactive Python Streamlit web dashboard capable of querying 50,000+ ATM transaction records with query response latency under 2 seconds.", "Specific: Streamlit web app; Measurable: <2s response; Achievable: indexed SQLite; Relevant: fleet monitoring; Time-bound: Milestones 1-15."),
        ("OBJ-02", "Establish a standardized 4-tier banking failure classification engine (Hardware, Network/Switch, Customer/Account, Host Timeout) mapping 100% of synthetic error codes.", "Specific: Complete failure categorization; Measurable: 100% code coverage; Achievable: standard ISO 8583 mappings; Relevant: root cause analysis; Time-bound: Step 8-10."),
        ("OBJ-03", "Implement an automated anomaly detection module to flag outlier ATM terminals whose rolling 7-day failure rate exceeds the 5.0% SLA threshold.", "Specific: Flag outlier terminals; Measurable: SLA >5% trigger; Achievable: statistical rolling window; Relevant: field dispatch; Time-bound: Step 10-14."),
        ("OBJ-04", "Develop 6+ responsive Plotly data visualizations (heatmaps, failure trend lines, pareto charts, and geo-cluster breakdowns) to cut failure root-cause identification time by 70%.", "Specific: 6 interactive charts; Measurable: 6 charts, 70% time reduction; Achievable: Plotly library; Relevant: operator UX; Time-bound: Step 12-15."),
        ("OBJ-05", "Achieve >85% automated test coverage using pytest for database query methods and transformation logic, verifying zero unhandled exception states prior to cloud deployment.", "Specific: Automated test suite; Measurable: >85% code coverage; Achievable: pytest unit test fixtures; Relevant: QA reliability; Time-bound: Step 13-15.")
    ]
    
    for row_idx, row_data in enumerate(obj_data, start=1):
        for col_idx, text in enumerate(row_data):
            obj_table.rows[row_idx].cells[col_idx].paragraphs[0].text = text
            
    style_table(obj_table, [1.0, 3.2, 2.3])
    doc.add_paragraph()
    
    # SECTION 7: Scope Definition
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("7. Scope Definition (In-Scope vs. Out-of-Scope)")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    scope_table = doc.add_table(rows=6, cols=2)
    scope_headers = ["In-Scope (Deliverables & Features)", "Out-of-Scope (Explicit Boundaries)"]
    for i, h in enumerate(scope_headers):
        scope_table.rows[0].cells[i].paragraphs[0].text = h
        
    scope_data = [
        ("Generation of a realistic 50,000+ synthetic transactional dataset modeling multi-bank ATM events over a 90-day observation window.", "Direct connection to live commercial banking core switches (e.g., ISO 8583 / NDC raw wire protocols) due to banking regulatory/security policies."),
        ("Normalized SQLite relational schema storing transactions, terminals, error codes, and maintenance records with foreign key constraints.", "Storage or ingestion of actual Personally Identifiable Information (PII) such as live 16-digit Primary Account Numbers (PAN), PINs, or CVVs."),
        ("Interactive Streamlit dashboard featuring multi-criteria filtering (date range, bank, terminal ID, transaction type, error category).", "Automated physical dispatch of field engineers or real-time bi-directional integration with enterprise ticketing systems (e.g., ServiceNow/Jira APIs)."),
        ("Statistical analysis and Plotly visualizations for failure trends, terminal ranking, latency correlation, and peak failure heatmaps.", "Interbank settlement, fee calculation, interchange clearing reconciliation, or actual funds settlement across member banks."),
        ("Comprehensive automated pytest test suite and zero-configuration public deployment on Streamlit Community Cloud.", "Native mobile applications (iOS/Android) or proprietary hardware terminal firmware flashing.")
    ]
    
    for row_idx, row_data in enumerate(scope_data, start=1):
        for col_idx, text in enumerate(row_data):
            scope_table.rows[row_idx].cells[col_idx].paragraphs[0].text = text
            
    style_table(scope_table, [3.25, 3.25])
    doc.add_paragraph()
    
    # SECTION 8: Assumptions and Constraints
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("8. Assumptions and Constraints")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    doc.add_heading("8.1 Operational & Data Assumptions", level=2)
    doc.add_paragraph("• Synthetic Data Fidelity: The synthetic dataset accurately mirrors retail banking operational parameters, incorporating realistic error distributions based on industry benchmarks (approx. 85-90% success, 10-15% varied failures).")
    doc.add_paragraph("• Response Code Conformity: Error code nomenclature adheres to standard ISO 8583 and NPCI National Financial Switch (NFS) conventions (e.g., Code 00: Approval, 51: Insufficient funds, 55: Incorrect PIN, 91: Issuer or switch inoperative, 96: System malfunction).")
    doc.add_paragraph("• Geographical Distribution: ATM terminals are distributed across Metro, Urban, and Semi-Urban zones, reflecting varying telecom stability and customer transaction densities.")
    doc.add_paragraph("• Target SLA Benchmark: A transaction failure rate of 5.0% or above is defined as an operational SLA breach triggering high-priority escalation.")
    
    doc.add_heading("8.2 Technical & Academic Constraints", level=2)
    doc.add_paragraph("• Compute & Memory Boundary: The production build must run smoothly within the 1 GB RAM allocation of free-tier Streamlit Community Cloud without out-of-memory terminations.")
    doc.add_paragraph("• Zero-Server Database: Application must rely strictly on a local, zero-configuration SQLite database engine embedded within the repository for portability and deterministic grading.")
    doc.add_paragraph("• Technology Standard: Pure Python 3.10+ implementation utilizing standard data science libraries (Pandas, Plotly, Streamlit, SQLite3, pytest) with Tailwind-styled UI elements.")
    doc.add_paragraph("• Academic Compliance: All project phases must strictly fulfill the MCA 2nd Year project curriculum requirements of ABES Engineering College under mentor Mr. Dinesh Sereno.")
    doc.add_paragraph()
    
    # SECTION 9: Success Criteria
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("9. Success Criteria (Measurable)")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    p = doc.add_paragraph("Project success will be objectively measured against the following quantitative performance metrics:")
    
    sc_table = doc.add_table(rows=6, cols=4)
    sc_headers = ["Metric ID", "Evaluation Parameter", "Success Target", "Verification Method"]
    for i, h in enumerate(sc_headers):
        sc_table.rows[0].cells[i].paragraphs[0].text = h
        
    sc_data = [
        ("SC-01", "Classification Completeness", "100% of error codes mapped to standard categories", "Automated SQL query confirming 0 unclassified or NULL failure records."),
        ("SC-02", "Interactive Query Latency", "< 2.0 seconds dashboard render on 50k+ rows", "Browser developer network profiling and Streamlit timer benchmarks."),
        ("SC-03", "Diagnostic Workflow Efficiency", "< 3 user clicks to isolate worst-performing terminal", "Usability test protocol tracking click path from Fleet Overview to Terminal Detail."),
        ("SC-04", "Test Code Coverage", ">= 85% statement coverage across core analytical modules", "pytest-cov execution report with zero failing assertions."),
        ("SC-05", "Public Cloud Availability", "100% uptime on Streamlit Community Cloud during viva", "Successful live demonstration during ABES project review under Mr. Dinesh Sereno.")
    ]
    
    for row_idx, row_data in enumerate(sc_data, start=1):
        for col_idx, text in enumerate(row_data):
            sc_table.rows[row_idx].cells[col_idx].paragraphs[0].text = text
            
    style_table(sc_table, [0.8, 1.8, 2.3, 1.6])
    doc.add_paragraph()
    
    # SECTION 10: Operational Question Feasibility Check
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("10. Operational Question Data Feasibility Check")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = RGBColor(26, 54, 93)
    
    add_callout(
        doc,
        "FEASIBILITY VERIFICATION AUDIT: Every operational question (OQ-01 through OQ-10) has been cross-referenced "
        "against the planned synthetic relational schema (tables: `transactions`, `atm_terminals`, `error_codes`, `maintenance_logs`). "
        "All required dimensions and metrics exist in the data model. No question requires external untracked parameters.",
        title="VERIFICATION CONFIRMATION"
    )
    
    # Save document
    doc.save(output_path)
    print(f"Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_dir = "docs"
    os.makedirs(out_dir, exist_ok=True)
    build_step1_document(os.path.join(out_dir, "01_Project_Understanding.docx"))
    # Also save with requested alias
    build_step1_document(os.path.join(out_dir, "01_name_project_understa.docx"))
