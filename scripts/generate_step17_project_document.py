import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def cm_to_dxa(cm):
    return int(cm * 567)

def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_border(cell, color="BFBFBF", sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def style_abes_table(table, col_widths_dxa):
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tblPr = table._tbl.tblPr
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="9638" w:type="dxa"/>')
    tblPr.append(tblW)
    
    # Header row
    hdr_row = table.rows[0]
    hdr_trPr = hdr_row._tr.get_or_add_trPr()
    hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    hdr_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    
    for c_idx, cell in enumerate(hdr_row.cells):
        set_cell_background(cell, "A51C30")
        set_cell_border(cell, "BFBFBF")
        set_cell_margins(cell, 60, 60, 90, 90)
        cell.width = Inches(col_widths_dxa[c_idx] / 1440.0)
        tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths_dxa[c_idx]}" w:type="dxa"/>')
        cell._tc.get_or_add_tcPr().append(tcW)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.bold = True
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)
                
    # Body rows
    for row in table.rows[1:]:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            set_cell_border(cell, "BFBFBF")
            set_cell_margins(cell, 60, 60, 90, 90)
            cell.width = Inches(col_widths_dxa[c_idx] / 1440.0)
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths_dxa[c_idx]}" w:type="dxa"/>')
            cell._tc.get_or_add_tcPr().append(tcW)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = "Calibri"
                    r.font.size = Pt(9)
                    r.font.color.rgb = RGBColor(21, 21, 20)

def add_abes_header_footer(doc):
    for s in doc.sections:
        s.top_margin = Inches(2.0 / 2.54) # 2 cm
        s.bottom_margin = Inches(2.0 / 2.54)
        s.left_margin = Inches(2.0 / 2.54)
        s.right_margin = Inches(2.0 / 2.54)
        s.header_distance = Inches(1.25 / 2.54) # 1.25 cm
        s.footer_distance = Inches(1.25 / 2.54)
        
        # Header
        hdr = s.header
        p_hdr = hdr.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_hdr.paragraph_format.space_after = Pt(0)
        r_hdr = p_hdr.add_run("ABES Engineering College | Project Document | p096")
        r_hdr.font.name = "Calibri"
        r_hdr.font.size = Pt(8)
        r_hdr.font.color.rgb = RGBColor(127, 127, 127)
        
        # Footer
        ftr = s.footer
        p_ftr = ftr.paragraphs[0]
        p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ftr.paragraph_format.space_after = Pt(0)
        r_ftr = p_ftr.add_run("Lavi Tewari | 2500320140085     Page ")
        r_ftr.font.name = "Calibri"
        r_ftr.font.size = Pt(8)
        r_ftr.font.color.rgb = RGBColor(127, 127, 127)
        
        # Add PAGE field
        fld_page = OxmlElement('w:fldSimple')
        fld_page.set(qn('w:instr'), 'PAGE')
        p_ftr._p.append(fld_page)

def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Georgia"
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(165, 28, 48) # A51C30
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Georgia"
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(21, 21, 20) # 151514
    return p

def add_body_p(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(21, 21, 20)
    return p

def add_note_p(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    r.italic = True
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(127, 127, 127)
    return p

def add_bullet_p(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(540 / 1440.0)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(21, 21, 20)
    return p

def build_project_document(output_path):
    doc = Document()
    add_abes_header_footer(doc)
    
    # ── PAGE 1: COVER ──
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(5)
    r1 = p1.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
    r1.bold = True
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(165, 28, 48)
    
    p_logo = doc.add_paragraph()
    p_logo.paragraph_format.space_after = Pt(5)
    r_logo = p_logo.add_run("[Insert ABES logo here]")
    r_logo.italic = True
    r_logo.font.name = "Calibri"
    r_logo.font.size = Pt(9)
    r_logo.font.color.rgb = RGBColor(127, 127, 127)
    
    p_title_lead = doc.add_paragraph()
    p_title_lead.paragraph_format.space_before = Pt(70)
    p_title_lead.paragraph_format.space_after = Pt(6)
    r_lead = p_title_lead.add_run("PROJECT DOCUMENT")
    r_lead.bold = True
    r_lead.font.name = "Georgia"
    r_lead.font.size = Pt(26)
    r_lead.font.color.rgb = RGBColor(21, 21, 20)
    
    p_main_title = doc.add_paragraph()
    p_main_title.paragraph_format.space_after = Pt(12)
    r_mt = p_main_title.add_run("p096 - ATM Transaction Failure Analysis")
    r_mt.bold = True
    r_mt.font.name = "Georgia"
    r_mt.font.size = Pt(18)
    r_mt.font.color.rgb = RGBColor(165, 28, 48)
    
    # Bottom border for title paragraph
    pBdr = parse_xml(f'''
        <w:pBdr {nsdecls("w")}>
            <w:bottom w:val="single" w:sz="12" w:space="4" w:color="A51C30"/>
        </w:pBdr>
    ''')
    p_main_title._p.get_or_add_pPr().append(pBdr)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(12)
    p_sub.paragraph_format.space_after = Pt(24)
    r_sub = p_sub.add_run("Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(127, 127, 127)
    
    # Cover Table: Field 2900 | Details 6738
    cov_tbl = doc.add_table(rows=8, cols=2)
    cov_widths = [2900, 6738]
    cov_data = [
        ("Project ID & Title", "p096 - ATM Transaction Failure Analysis"),
        ("Domain", "Banking & Retail Payment Systems / Data Analytics"),
        ("Programme / Batch", "MCA 2nd Year Data Analyst - ABES_OND26_MCA Batch 01"),
        ("Learner(s) & Roll No(s)", "Lavi Tewari (Roll No: 2500320140085)"),
        ("Mentor", "Mr. Dinesh Sereno"),
        ("GitHub repository", "https://github.com/Lavitewari11/atmTransaction"),
        ("Live app (Streamlit)", "https://2500320140085-p062-atm.streamlit.app/"),
        ("Version & date", "v1.0 - 21/11/2026")
    ]
    for r_idx, (fld, det) in enumerate(cov_data):
        cov_tbl.rows[r_idx].cells[0].paragraphs[0].text = fld
        cov_tbl.rows[r_idx].cells[1].paragraphs[0].text = det
    style_abes_table(cov_tbl, cov_widths)
    
    doc.add_page_break()
    
    # ── PAGE 2: TABLE OF CONTENTS ──
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_after = Pt(12)
    fld_toc = parse_xml(f'''
        <w:fldSimple {nsdecls("w")} w:instr="TOC \\o &quot;1-1&quot; \\h \\z \\u"/>
    ''')
    p_toc._p.append(fld_toc)
    
    # Insert initial TOC text
    toc_items = [
        "1. Project Description", "2. How Might We (HMW) Statements", "3. Stakeholders",
        "4. Traceability Matrix", "5. User Stories", "6. Business Requirements Document (BRD)",
        "7. Technical Requirements Document (TRD)", "8. Database Design", "9. Test Cases",
        "10. UI Framework", "Appendix A - Prompt to generate this document", "Appendix B - AI tools disclosure"
    ]
    for ti in toc_items:
        p_ti = doc.add_paragraph()
        p_ti.paragraph_format.space_after = Pt(2)
        r_ti = p_ti.add_run(ti)
        r_ti.font.name = "Calibri"
        r_ti.font.size = Pt(10)
        
    doc.add_page_break()
    
    # ── SECTION 1: PROJECT DESCRIPTION ──
    add_h1(doc, "1. Project Description")
    p1_tbl = doc.add_table(rows=7, cols=2)
    p1_widths = [2400, 7238]
    p1_data = [
        ("Overview", "Developing an end-to-end interactive diagnostic data analytics platform for ATM terminal networks, analyzing failure telemetry across 50,003 records."),
        ("Problem statement", "ATM terminal downtime and transaction failures (technical timeouts, cash dispenser jams, switch link drops, and customer authentication errors) reduce success rates below banking SLAs, erode customer trust, and inflate maintenance costs."),
        ("Intended users", "ATM Operations Managers (daily fleet health), Field Maintenance Leads (hardware repair), NOC Analysts (telecom/switch timeouts), Customer Dispute Officers (chargebacks)."),
        ("In scope", "Synthetic 50,003 multi-bank dataset (90 days), normalized SQLite schema (4 tables + 1 view), EDA, Pareto root-cause analysis, rolling SLA outlier detection (>5%), MTBF metrics, Tailwind UI, pytest test suite, Streamlit Cloud deployment."),
        ("Out of scope", "Direct ISO 8583 live switch connection, ingestion of actual PII (16-digit PAN, PIN, CVV), automated physical field engineer dispatch, and interbank monetary settlement."),
        ("Assumptions", "Synthetic transaction dataset mirrors retail banking failure distributions (~10-15% failures); ISO 8583 response codes standard; SQLite sufficient for read-intensive analytics; 1 GB RAM cloud ceiling."),
        ("Dataset", "Synthetic multi-bank dataset: 50,003 transaction rows across 200 terminals, 30 standard ISO error codes, and 1,500 maintenance events with 5 embedded operational patterns.")
    ]
    for r_idx, (it, desc) in enumerate(p1_data):
        p1_tbl.rows[r_idx].cells[0].paragraphs[0].text = it
        p1_tbl.rows[r_idx].cells[1].paragraphs[0].text = desc
    style_abes_table(p1_tbl, p1_widths)
    
    # ── SECTION 2: HMW STATEMENTS ──
    add_h1(doc, "2. How Might We (HMW) Statements")
    add_note_p(doc, "NOTE Format: How might we <action> for <user> so that <outcome>?")
    hmw_tbl = doc.add_table(rows=11, cols=2)
    hmw_widths = [1300, 8338]
    hmws = [
        ("HMW ID", "HMW statement"),
        ("HMW-01", "How might we provide fleet operations with instant visibility into real-time failure rate trends across hourly, daily, and weekly cycles?"),
        ("HMW-02", "How might we systematically categorize transaction declines into clear banking root causes (Hardware, Network, Customer, Host Timeout)?"),
        ("HMW-03", "How might we automatically pinpoint and rank chronic outlier terminals whose failure rates persistently breach the 5% SLA threshold?"),
        ("HMW-04", "How might we visualize temporal peaks in network timeouts to help NOC teams anticipate peak switch congestion?"),
        ("HMW-05", "How might we expose failure rate variations across different transaction operations (Withdrawal vs. Inquiry vs. PIN Change)?"),
        ("HMW-06", "How might we correlate transaction round-trip response latency with timeout likelihood to establish proactive network warnings?"),
        ("HMW-07", "How might we isolate transaction failure behavior between On-Us and Off-Us networks to pinpoint external inter-bank switch bottlenecks?"),
        ("HMW-08", "How might we quantify the total financial transaction value disrupted by terminal hardware faults versus customer balance errors?"),
        ("HMW-09", "How might we benchmark component hardware reliability across different ATM manufacturers and terminal models?"),
        ("HMW-10", "How might we calculate the Mean Time Between Failures (MTBF) and recurrent fault cadence for unstable terminals?")
    ]
    for r_idx, (hid, htext) in enumerate(hmws):
        hmw_tbl.rows[r_idx].cells[0].paragraphs[0].text = hid
        hmw_tbl.rows[r_idx].cells[1].paragraphs[0].text = htext
    style_abes_table(hmw_tbl, hmw_widths)
    
    # ── SECTION 3: STAKEHOLDERS ──
    add_h1(doc, "3. Stakeholders")
    sh_tbl = doc.add_table(rows=7, cols=5)
    sh_widths = [1000, 2400, 1900, 3138, 1200]
    shs = [
        ("ID", "Stakeholder", "Role", "Interest", "Influence"),
        ("SH-01", "ATM Operations Manager", "Fleet Operations Lead", "Fleet uptime, availability SLAs, cash cycle planning", "High"),
        ("SH-02", "Field Maintenance Lead", "Hardware Support Lead", "Dispatches technicians, orders parts, schedules PM", "Medium"),
        ("SH-03", "NOC & Switch Network Analyst", "Switch Engineer", "Monitors switch timeouts, gateway latency, network links", "Medium"),
        ("SH-04", "Customer Dispute Officer", "Customer Grievance", "Verifies cash dispense failures vs. claims, resolves chargebacks", "Low"),
        ("SH-05", "Chief Operations Officer (COO)", "Executive Sponsor", "Allocates budget, approves vendor contracts, reviews macro SLA", "High"),
        ("SH-06", "Academic Project Mentor", "Faculty Guide (Mr. Dinesh Sereno)", "Evaluates analytical methodology, software hygiene, curriculum", "High")
    ]
    for r_idx, r_data in enumerate(shs):
        for c_idx, val in enumerate(r_data):
            sh_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(sh_tbl, sh_widths)
    
    # ── SECTION 4: TRACEABILITY MATRIX ──
    add_h1(doc, "4. Traceability Matrix")
    add_note_p(doc, "NOTE One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")
    rtm_tbl = doc.add_table(rows=13, cols=9)
    rtm_widths = [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040]
    rtms = [
        ("HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"),
        ("HMW-01", "BR-01", "Fleet failure rate trending", "SH-01", "FR-01, 04", "Must", "US-01, 04", "TC-F-01, 04", "UI-01"),
        ("HMW-02", "BR-02", "4-tier failure classification", "SH-01, 03", "FR-02, 03", "Must", "US-02, 03", "TC-F-02, 07", "UI-02"),
        ("HMW-03", "BR-03", "Chronic outlier detection (>5%)", "SH-01, 02", "FR-05", "Must", "US-10", "TC-F-10", "UI-03"),
        ("HMW-04", "BR-04", "24x7 temporal failure heatmap", "SH-03", "FR-06", "Should", "US-08", "TC-F-08", "UI-04"),
        ("HMW-05", "BR-05", "Transaction-type segmentation", "SH-01, 04", "FR-07", "Must", "US-07, 09", "TC-F-09", "UI-02"),
        ("HMW-06", "BR-06", "Latency vs. timeout correlation", "SH-03", "FR-08, 09", "Must", "US-13", "TC-F-13", "UI-05"),
        ("HMW-07", "BR-07", "On-Us vs Off-Us segregation", "SH-01, 03", "FR-10", "Should", "US-14", "TC-F-14", "UI-05"),
        ("HMW-08", "BR-08", "Financial value disruption sum", "SH-01, 05", "FR-11", "Must", "US-06", "TC-F-06", "UI-06"),
        ("HMW-09", "BR-09", "Hardware vendor benchmarking", "SH-02, 05", "FR-12", "Should", "US-12", "TC-F-12", "UI-03"),
        ("HMW-10", "BR-10", "MTBF predictive metric", "SH-02, 06", "FR-13", "Must", "US-11", "TC-F-11", "UI-03"),
        ("HMW-01", "BR-11", "Dynamic multi-criteria filter", "SH-01, 02", "FR-15", "Could", "US-01, 05", "TC-F-05", "UI-01..06"),
        ("HMW-03", "BR-12", "Filtered CSV / Excel export", "SH-02, 04", "FR-16", "Could", "US-15", "TC-F-15", "UI-08")
    ]
    for r_idx, r_data in enumerate(rtms):
        for c_idx, val in enumerate(r_data):
            rtm_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(rtm_tbl, rtm_widths)
    
    # ── SECTION 5: USER STORIES ──
    add_h1(doc, "5. User Stories")
    us_tbl = doc.add_table(rows=16, cols=4)
    us_widths = [900, 3400, 4138, 1200]
    uss = [
        ("US-ID", "User story", "Acceptance criteria (Given / When / Then)", "Priority"),
        ("US-01", "As an ATM Ops Manager, I want cached 50,000 records with date filtering so that I view specific periods.", "Given app loads, When date selected, Then records filter in < 2.0s without re-query.", "Must"),
        ("US-02", "As a NOC Analyst, I want error codes mapped to 4-tier taxonomy so that 100% of errors are classified.", "Given raw errors, When classified, Then 0 NULL categories remain.", "Must"),
        ("US-03", "As a Field Lead, I want a data quality audit tile so that I can inspect quarantined records.", "Given anomalies, When app opens, Then sidebar tile displays exact counts.", "Should"),
        ("US-04", "As an ATM Ops Manager, I want fleet failure rate KPI and trend curve so that I spot intraday surges.", "Given records loaded, When Fleet Overview opens, Then KPI & line chart render < 1.5s.", "Must"),
        ("US-05", "As an ATM Ops Manager, I want multi-criteria sidebar filters so that I isolate specific bank/zone segments.", "Given dropdowns selected, When changed, Then state syncs across all pages.", "Could"),
        ("US-06", "As a COO, I want an executive summary row so that I evaluate business INR at risk.", "Given executive access, When viewing overview, Then 4 KPI tiles show volume & INR loss.", "Must"),
        ("US-07", "As a NOC Analyst, I want an 80/20 Pareto chart so that I target the vital few failure categories.", "Given failures aggregated, When Pareto renders, Then bars sort descending with 80% line.", "Must"),
        ("US-08", "As a NOC Analyst, I want a 24x7 failure heatmap so that I anticipate peak switch congestion.", "Given temporal data, When heatmap renders, Then weekend peak hours clearly highlighted.", "Should"),
        ("US-09", "As a Dispute Officer, I want failure rates by transaction type so that I prioritize cash dispense claims.", "Given disputes, When type chart renders, Then Cash Withdrawal vs Inquiry is clear.", "Must"),
        ("US-10", "As a Field Lead, I want outlier terminals ranked (>5% SLA) so that I dispatch engineers to chronic nodes.", "Given rolling 7-day rate, When >5%, Then outlier terminals show red alert badges.", "Must"),
        ("US-11", "As a Field Lead, I want MTBF computed per terminal so that I schedule preventive servicing.", "Given failure timestamps, When calculating MTBF, Then operating hours display correctly.", "Must"),
        ("US-12", "As a COO, I want vendor hardware failure rates benchmarked so that I enforce warranty penalties.", "Given vendor telemetry, When benchmark renders, Then failure rate per 1,000 txns shows.", "Should"),
        ("US-13", "As a NOC Analyst, I want latency scatter with Pearson r so that I prove latency causes timeouts.", "Given latency logged, When scatter renders, Then latency > 3500ms shows 100% failure.", "Must"),
        ("US-14", "As an ATM Ops Manager, I want On-Us vs Off-Us Z-test so that I prove interbank switch bottlenecks.", "Given card types, When compared, Then donut + Z-test prove statistical significance.", "Should"),
        ("US-15", "As a Field Lead, I want to download filtered CSV so that I have offline diagnostic sheets.", "Given filtered table, When download clicked, Then CSV file downloads in < 2.0s.", "Could")
    ]
    for r_idx, r_data in enumerate(uss):
        for c_idx, val in enumerate(r_data):
            us_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(us_tbl, us_widths)
    
    # ── SECTION 6: BRD ──
    add_h1(doc, "6. Business Requirements Document (BRD)")
    add_h2(doc, "6.1 Business objectives")
    add_bullet_p(doc, "OBJ-01: Build an interactive Python Streamlit analytics dashboard capable of querying 50,000+ ATM records in < 2.0s.")
    add_bullet_p(doc, "OBJ-02: Establish a standardized 4-tier banking failure classification engine with 100% categorical mapping.")
    add_bullet_p(doc, "OBJ-03: Implement automated anomaly detection to flag outlier terminals exceeding the 5.0% rolling failure SLA threshold.")
    add_bullet_p(doc, "OBJ-04: Deliver 6+ responsive Plotly visualizations to cut failure root-cause identification time by 70%.")
    add_bullet_p(doc, "OBJ-05: Achieve >85% automated unit test coverage with pytest across core analytical modules prior to deployment.")
    
    add_h2(doc, "6.2 Success KPIs")
    kpi_tbl = doc.add_table(rows=6, cols=5)
    kpi_widths = [1900, 2100, 2100, 1700, 1838]
    kpis = [
        ("KPI", "Numerator", "Denominator", "Period", "Target / note"),
        ("Fleet Failure Rate", "Failed Transactions", "Total Transactions", "Hourly / Daily", "< 5.0% SLA Target"),
        ("Terminal Outlier Rate", "Terminal Failed Txns", "Terminal Total Txns", "Rolling 7-Day", "> 5.0% triggers alert"),
        ("Mean Time Between Failures", "Operating Hours - Downtime", "Total Failure Count", "90-Day Window", "> 48.0 Hours"),
        ("Interbank Failure Ratio", "Off-Us Failure Rate", "On-Us Failure Rate", "Monthly", "< 1.2x Target"),
        ("High Latency Failure Rate", "Failed Txns > 3500ms", "Total Txns > 3500ms", "Continuous", "< 10.0% Target")
    ]
    for r_idx, r_data in enumerate(kpis):
        for c_idx, val in enumerate(r_data):
            kpi_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(kpi_tbl, kpi_widths)
    
    add_h2(doc, "6.3 Business requirements")
    br_tbl = doc.add_table(rows=13, cols=4)
    br_widths = [1100, 5338, 1300, 1900]
    brs = [
        ("BR-ID", "Requirement", "Priority", "Source (stakeholder)"),
        ("BR-01", "The business needs to aggregate and track fleet failure rates over hourly/daily windows.", "Must", "SH-01 (ATM Ops Manager)"),
        ("BR-02", "The business needs to map raw error codes into a 4-tier banking failure taxonomy.", "Must", "SH-01, 03 (NOC Analyst)"),
        ("BR-03", "The business needs to automatically identify and rank terminals exceeding 5.0% failure SLA.", "Must", "SH-01, 02 (Field Lead)"),
        ("BR-04", "The business needs to visualize temporal failure density across a 24x7 heatmap.", "Should", "SH-03 (NOC Analyst)"),
        ("BR-05", "The business needs to segment failure volumes across transaction operations.", "Must", "SH-01, 04 (Dispute Officer)"),
        ("BR-06", "The business needs to correlate network round-trip response latency with timeout frequency.", "Must", "SH-03 (NOC Analyst)"),
        ("BR-07", "The business needs to segregate failure metrics between On-Us and Off-Us card networks.", "Should", "SH-01, 03 (Switch Analyst)"),
        ("BR-08", "The business needs to calculate gross financial transaction value disrupted by terminal faults.", "Must", "SH-01, 05 (COO)"),
        ("BR-09", "The business needs to benchmark hardware failure frequencies across ATM manufacturers.", "Should", "SH-02, 05 (COO)"),
        ("BR-10", "The business needs to compute MTBF in running hours for unstable terminals.", "Must", "SH-02, 06 (Mentor)"),
        ("BR-11", "The business needs dynamic multi-criteria sidebar filters across all dashboard views.", "Could", "SH-01, 02 (Field Lead)"),
        ("BR-12", "The business needs to export filtered transaction records and audit reports as CSV.", "Could", "SH-02, 04 (Dispute Officer)")
    ]
    for r_idx, r_data in enumerate(brs):
        for c_idx, val in enumerate(r_data):
            br_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(br_tbl, br_widths)
    
    add_h2(doc, "6.4 Constraints and out of scope")
    add_bullet_p(doc, "Direct live integration with commercial banking core switches (ISO 8583 / NDC wire protocols) is prohibited due to regulatory mandates.")
    add_bullet_p(doc, "Storage or handling of real customer Personally Identifiable Information (PII) such as live PANs or PINs is strictly out of scope.")
    add_bullet_p(doc, "Automated field engineer dispatch APIs or integration with enterprise ticketing systems (e.g. ServiceNow) is deferred to future enterprise phases.")
    add_bullet_p(doc, "All execution must operate within free-tier Streamlit Community Cloud compute boundaries (1 GB RAM allocation).")
    
    # ── SECTION 7: TRD ──
    add_h1(doc, "7. Technical Requirements Document (TRD)")
    add_h2(doc, "7.1 Technology stack (all free)")
    stack_tbl = doc.add_table(rows=9, cols=3)
    stack_widths = [2000, 3800, 3838]
    stacks = [
        ("Layer", "Tool", "Why"),
        ("Language", "Python 3.11 / 3.12", "Standard programming language for data analytics and scientific computing."),
        ("Data processing", "Pandas 2.2+, NumPy 1.26+", "High-performance vectorized tabular processing, slicing, and aggregations."),
        ("Storage", "SQLite 3.45+ (data/app.db)", "Zero-configuration, serverless, self-contained embedded relational database."),
        ("Charts", "Plotly 5.22+ (Plotly Express)", "Interactive web-native charts, heatmaps, hover tooltips, and Pareto lines."),
        ("App / UI", "Streamlit 1.37+ & Tailwind CSS", "Rapid multipage data application framework with custom Tailwind styling."),
        ("Testing", "pytest 8.2+, pytest-cov", "Industry-standard automated test runner with statement coverage reports."),
        ("Version control", "Git & GitHub", "Source code versioning, collaboration, and continuous cloud deployment."),
        ("Editor", "VS Code / Antigravity IDE", "Rich developer tooling, linting, interactive debugging, and shell execution.")
    ]
    for r_idx, r_data in enumerate(stacks):
        for c_idx, val in enumerate(r_data):
            stack_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(stack_tbl, stack_widths)
    
    add_h2(doc, "7.2 Functional requirements")
    fr_tbl = doc.add_table(rows=17, cols=3)
    fr_widths = [1100, 6338, 2200]
    frs = [
        ("FR-ID", "Requirement", "Linked BR / US"),
        ("FR-01", "Load SQLite tables into cached DataFrames using @st.cache_data.", "BR-01 / US-01"),
        ("FR-02", "Validate schema and quarantine corrupted records on startup.", "BR-02 / US-02, 03"),
        ("FR-03", "Map raw error codes to 4-tier failure taxonomy via lookup join.", "BR-02 / US-02"),
        ("FR-04", "Compute fleet failure rates aggregated by hourly and daily windows.", "BR-01 / US-04, 06"),
        ("FR-05", "Flag chronic outlier terminals exceeding rolling 5.0% failure SLA.", "BR-03 / US-10"),
        ("FR-06", "Generate 24-hour x 7-day temporal failure density heatmap matrix.", "BR-04 / US-08"),
        ("FR-07", "Segment transaction failure distributions by operation type.", "BR-05 / US-07, 09"),
        ("FR-08", "Calculate network round-trip response latency percentiles (p50, p90, p99).", "BR-06 / US-13"),
        ("FR-09", "Compute Pearson correlation between latency and failure outcome.", "BR-06 / US-13"),
        ("FR-10", "Segregate failure volume and calculate On-Us vs Off-Us Z-test.", "BR-07 / US-14"),
        ("FR-11", "Aggregate gross INR transaction value disrupted by failure category.", "BR-08 / US-06"),
        ("FR-12", "Benchmark vendor hardware failure rates per 1,000 transactions.", "BR-09 / US-12"),
        ("FR-13", "Compute MTBF in operating hours per terminal.", "BR-10 / US-11"),
        ("FR-14", "Render standardized Plotly figures with shared responsive theme.", "BR-01..10 / US-04..14"),
        ("FR-15", "Synchronize global sidebar filter state across multipage views.", "BR-11 / US-01, 05"),
        ("FR-16", "Stream filtered transaction extracts to CSV via download button.", "BR-12 / US-15")
    ]
    for r_idx, r_data in enumerate(frs):
        for c_idx, val in enumerate(r_data):
            fr_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(fr_tbl, fr_widths)
    
    add_h2(doc, "7.3 Non-functional requirements")
    nfr_tbl = doc.add_table(rows=6, cols=2)
    nfr_widths = [1100, 8538]
    nfrs = [
        ("NFR-ID", "Requirement"),
        ("NFR-01", "Performance: Page rendering latency must execute in < 2.0 seconds on 50,000+ cached records."),
        ("NFR-02", "Reliability: Automated unit test suite must achieve >= 85% statement coverage with zero test failures."),
        ("NFR-03", "Security: SQLite database must open in read-only URI mode (mode=ro) preventing accidental mutation."),
        ("NFR-04", "Privacy: Zero live Personally Identifiable Information (PAN, PIN, CVV) in repository."),
        ("NFR-05", "Portability: Pinned requirements.txt ensuring deterministic execution on Linux, Windows, and macOS.")
    ]
    for r_idx, r_data in enumerate(nfrs):
        for c_idx, val in enumerate(r_data):
            nfr_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(nfr_tbl, nfr_widths)
    
    add_h2(doc, "7.4 Architecture and deployment")
    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.space_after = Pt(5)
    r_arch = p_arch.add_run("Raw data (data/raw) -> <clean script> -> <SQLite or xlsx> (data/processed) -> Streamlit app (app.py) -> user")
    r_arch.bold = True
    r_arch.font.name = "Calibri"
    r_arch.font.size = Pt(10)
    r_arch.font.color.rgb = RGBColor(165, 28, 48)
    
    dep_tbl = doc.add_table(rows=4, cols=3)
    dep_widths = [2200, 4838, 2600]
    deps = [
        ("Stage", "Command / action", "Done when"),
        ("1. Local", "streamlit run app.py & pytest -v", "All 11 tests pass, dashboard renders locally"),
        ("2. GitHub", "git push origin main", "Repository updated with clean commit history"),
        ("3. Streamlit Community Cloud", "share.streamlit.io deploy", "App live on public URL with zero launch errors")
    ]
    for r_idx, r_data in enumerate(deps):
        for c_idx, val in enumerate(r_data):
            dep_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(dep_tbl, dep_widths)
    
    # ── SECTION 8: DATABASE DESIGN ──
    add_h1(doc, "8. Database Design")
    add_note_p(doc, "NOTE Storage used: SQLite (data/app.db). Raw data stays unchanged; cleaned data is stored here.")
    db_tbl = doc.add_table(rows=17, cols=5)
    db_widths = [1700, 2300, 2100, 1400, 2138]
    dbs = [
        ("Table / sheet", "Column", "Type", "Key", "Description"),
        ("atm_terminals", "terminal_id", "TEXT", "PK", "Unique terminal identifier (ATM_001..200)"),
        ("atm_terminals", "bank_owner", "TEXT", "—", "Acquiring bank institution"),
        ("atm_terminals", "vendor", "TEXT", "—", "Hardware manufacturer (NCR, Diebold, etc.)"),
        ("atm_terminals", "city", "TEXT", "—", "Geographic operating city"),
        ("atm_terminals", "zone", "TEXT", "—", "Classification: Metro, Urban, Semi-Urban"),
        ("error_codes", "error_code", "TEXT", "PK", "ISO 8583 banking response code"),
        ("error_codes", "failure_category", "TEXT", "—", "4-tier category: Hardware, Network, etc."),
        ("error_codes", "severity", "TEXT", "—", "Impact level: Critical, High, Medium, Low"),
        ("transactions", "transaction_id", "TEXT", "PK", "Unique transaction event token"),
        ("transactions", "terminal_id", "TEXT", "FK", "Reference to atm_terminals"),
        ("transactions", "timestamp", "DATETIME", "—", "Transaction execution datetime"),
        ("transactions", "transaction_type", "TEXT", "—", "Withdrawal, Inquiry, PIN Change"),
        ("transactions", "amount", "REAL", "—", "Financial value in INR"),
        ("transactions", "status", "TEXT", "—", "Outcome: Approved, Failed, Reversed"),
        ("transactions", "error_code", "TEXT", "FK", "Reference to error_codes"),
        ("transactions", "response_time_ms", "INTEGER", "—", "Network round-trip latency in ms")
    ]
    for r_idx, r_data in enumerate(dbs):
        for c_idx, val in enumerate(r_data):
            db_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(db_tbl, db_widths)
    
    p_rel = doc.add_paragraph()
    p_rel.paragraph_format.space_before = Pt(5)
    r_rel = p_rel.add_run("Relationship: atm_terminals (1) -- (many) transactions | error_codes (1) -- (many) transactions | atm_terminals (1) -- (many) maintenance_logs")
    r_rel.bold = True
    r_rel.font.name = "Calibri"
    r_rel.font.size = Pt(9.5)
    
    p_ddl_lbl = doc.add_paragraph()
    p_ddl_lbl.paragraph_format.space_before = Pt(5)
    r_ddl_lbl = p_ddl_lbl.add_run("SQLite DDL:")
    r_ddl_lbl.bold = True
    r_ddl_lbl.font.name = "Calibri"
    r_ddl_lbl.font.size = Pt(9.5)
    
    p_ddl_code = doc.add_paragraph()
    r_ddl_code = p_ddl_code.add_run(
        "CREATE TABLE atm_terminals (terminal_id TEXT PRIMARY KEY, bank_owner TEXT, vendor TEXT, city TEXT, zone TEXT);\n"
        "CREATE TABLE error_codes (error_code TEXT PRIMARY KEY, failure_category TEXT, severity TEXT);\n"
        "CREATE TABLE transactions (transaction_id TEXT PRIMARY KEY, terminal_id TEXT REFERENCES atm_terminals, timestamp DATETIME, amount REAL, status TEXT, error_code TEXT REFERENCES error_codes, response_time_ms INTEGER);"
    )
    r_ddl_code.font.name = "Consolas"
    r_ddl_code.font.size = Pt(9)
    
    # ── SECTION 9: TEST CASES ──
    add_h1(doc, "9. Test Cases")
    tc_tbl = doc.add_table(rows=12, cols=6)
    tc_widths = [900, 1500, 2400, 2638, 1100, 1100]
    tcs = [
        ("TC-ID", "Linked Req / US", "Steps", "Expected result", "Actual", "Pass / Fail"),
        ("TC-T-01", "BR-01 / US-01", "Open SQLite connection in mode=ro.", "Connects successfully without lock.", "Connects valid", "Pass"),
        ("TC-T-02", "BR-01 / US-01", "Check table counts (terminals>=200, txns>=50k).", "Minimum rows present.", "200 & 50,003", "Pass"),
        ("TC-T-03", "BR-02 / US-02", "Assert primary key uniqueness on masters.", "Zero duplicate keys.", "0 duplicates", "Pass"),
        ("TC-T-04", "BR-02 / US-02", "Assert zero orphan FK rows in transactions.", "Zero orphan records.", "0 orphans", "Pass"),
        ("TC-T-05", "BR-01 / US-01", "Verify v_transaction_analysis view column count.", "View contains >= 30 columns.", "35 columns", "Pass"),
        ("TC-T-06", "BR-01 / US-04", "Test failure rate math on known fixture array.", "Exact 40.0% calculated.", "40.0% exact", "Pass"),
        ("TC-T-07", "BR-03 / US-10", "Test outlier threshold filter trigger > 5.0%.", "Only ATM_001 flagged.", "ATM_001 only", "Pass"),
        ("TC-T-08", "BR-10 / US-11", "Test MTBF formula on known values.", "Exact 40.0 hours calculated.", "40.0 hrs", "Pass"),
        ("TC-T-09", "BR-08 / US-06", "Test financial disruption SUM on Failed rows.", "Exact ₹12,000 aggregated.", "₹12,000", "Pass"),
        ("TC-T-10", "BR-11 / US-05", "Verify app.py and all 6 page files exist.", "All 8 Python files present.", "All present", "Pass"),
        ("TC-T-11", "BR-11 / US-05", "Verify .streamlit/config.toml primaryColor.", "Config primaryColor is #2B6CB0.", "#2B6CB0 valid", "Pass")
    ]
    for r_idx, r_data in enumerate(tcs):
        for c_idx, val in enumerate(r_data):
            tc_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(tc_tbl, tc_widths)
    
    # ── SECTION 10: UI FRAMEWORK ──
    add_h1(doc, "10. UI Framework")
    add_note_p(doc, "NOTE The screen-by-screen design, built first as an HTML framework, then in Streamlit.")
    ui_tbl = doc.add_table(rows=9, cols=6)
    ui_widths = [800, 1200, 1700, 2900, 1700, 1338]
    uis = [
        ("UI-ID", "Screen", "Purpose", "Components", "Filters", "User stories"),
        ("UI-01", "Fleet Overview", "Macro fleet KPIs & failure trend", "4 KPI cards, hourly line chart, summary table", "Date, Bank, Zone", "US-04, US-06"),
        ("UI-02", "Failure Category", "Root-cause Pareto & type breakdown", "Pareto 80/20 chart, grouped bar, donut split", "Date, Bank, Cat", "US-07, US-09"),
        ("UI-03", "Terminal Health", "Outlier ranking & vendor reliability", "SLA table (>5%), MTBF bar, vendor benchmark", "Date, Bank, Zone", "US-10..12"),
        ("UI-04", "Temporal Patterns", "24x7 failure congestion heatmap", "Plotly density heatmap, peak callout tile", "Date, Bank, Cat", "US-08"),
        ("UI-05", "Network & Latency", "Latency correlation & switch share", "Scatter plot, Pearson r metric, donut chart", "Date, Bank, Card", "US-13, US-14"),
        ("UI-06", "Financial Impact", "Quantify disrupted transaction INR", "Disrupted value KPI cards, treemap chart", "Date, Bank, Type", "US-06"),
        ("UI-07", "Data Quality", "Audit missing/quarantined data", "Sidebar audit tile with exact anomaly counts", "Global", "US-03"),
        ("UI-08", "Data Export", "Offline inspection sheet extraction", "st.download_button, CSV stream generator", "All Filters", "US-15")
    ]
    for r_idx, r_data in enumerate(uis):
        for c_idx, val in enumerate(r_data):
            ui_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(ui_tbl, ui_widths)
    
    doc.add_page_break()
    
    # ── APPENDIX A: PROMPT ──
    add_h1(doc, "Appendix A - Prompt to generate this document")
    p_pref_lbl = doc.add_paragraph()
    r_pref_lbl = p_pref_lbl.add_run("MY PREFERENCES")
    r_pref_lbl.bold = True
    r_pref_lbl.font.name = "Calibri"
    r_pref_lbl.font.size = Pt(10)
    r_pref_lbl.font.color.rgb = RGBColor(165, 28, 48)
    
    p_pref_code = doc.add_paragraph()
    r_pref_code = p_pref_code.add_run(
        "PROJECT ID & TITLE: p096 - ATM Transaction Failure Analysis\n"
        "DOMAIN: Data Analytics\n"
        "LEARNER(S) & ROLL NO(S): Lavi Tewari, 2500320140085\n"
        "MENTOR: Mr. Dinesh Sereno\n"
        "BATCH: MCA 2nd Year Data Analyst - ABES_OND26_MCA Batch 01\n"
        "INTENDED USERS: ATM Fleet Managers, NOC Engineers, Field Service Leads, Dispute Officers\n"
        "PROBLEM IN MY WORDS: ATM failures cause SLA drops and customer friction without centralized visibility.\n"
        "DATASET: Synthetic multi-bank ATM operational logs (50,003 records, 90 days)\n"
        "KPIS I WANT: Fleet Failure Rate, Outlier Terminals (>5%), MTBF, Latency Correlation, INR Disrupted\n"
        "FEATURES I WANT: Tailwind KPIs, Pareto chart, Heatmap, Scatter plot, Treemap, CSV Export\n"
        "STORAGE: SQLite embedded database (data/app.db)\n"
        "TEAM SIZE & DEADLINE: 1 Developer, Academic Milestone Evaluation"
    )
    r_pref_code.font.name = "Consolas"
    r_pref_code.font.size = Pt(9)
    
    p_inst_lbl = doc.add_paragraph()
    p_inst_lbl.paragraph_format.space_before = Pt(10)
    r_inst_lbl = p_inst_lbl.add_run("INSTRUCTIONS FOR THE AI")
    r_inst_lbl.bold = True
    r_inst_lbl.font.name = "Calibri"
    r_inst_lbl.font.size = Pt(10)
    r_inst_lbl.font.color.rgb = RGBColor(165, 28, 48)
    
    p_inst_code = doc.add_paragraph()
    r_inst_code = p_inst_code.add_run(
        "Build ONE Data Analyst project end to end in 18 steps with full professional depth, "
        "retaining strict ID consistency (OQ, SH, HMW, BR, FR, US, UI, TC), using synthetic data, "
        "delivering Word/Excel deliverables, and generating final ABES Word report and slide deck."
    )
    r_inst_code.font.name = "Consolas"
    r_inst_code.font.size = Pt(9)
    
    # ── APPENDIX B: AI TOOLS DISCLOSURE ──
    add_h1(doc, "Appendix B - AI tools disclosure")
    ai_tbl = doc.add_table(rows=4, cols=3)
    ai_widths = [2600, 3500, 3538]
    ais = [
        ("Tool", "Used for", "How the output was checked"),
        ("Antigravity AI (Gemini 3.8 Flash / Claude)", "Assisting with requirements structuring, document templating, and test fixture scaffolding.", "Hand-audited against ABES rubrics, verified with pytest unit test executions."),
        ("Python python-docx", "Automated generation of Word project deliverables with exact dxa table layouts and A51C30 styling.", "Opened and visually inspected; validated cell borders, font sizes, and cantSplit tags."),
        ("Python python-pptx", "Automated generation of the 11-slide presentation deck adhering to the ABES canvas guidelines.", "Inspected slide geometry, card coordinates, font sizes, and speaker notes.")
    ]
    for r_idx, r_data in enumerate(ais):
        for c_idx, val in enumerate(r_data):
            ai_tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_abes_table(ai_tbl, ai_widths)
    
    doc.save(output_path)
    print(f"Step 17 ABES Project Document saved at: {output_path}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_project_document("docs/p096_Project_Document.docx")
