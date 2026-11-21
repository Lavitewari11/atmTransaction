import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ─────────────────────────── shared helpers ────────────────────────────
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
        set_shd(c, hbg); set_margins(c, 110, 110, 130, 130)
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.bold = True; r.font.color.rgb = RGBColor(255,255,255)
                r.font.size = Pt(9.5); r.font.name = "Calibri"
    for ri, row in enumerate(table.rows[1:], 1):
        bg = alt if ri % 2 else "FFFFFF"
        for c in row.cells:
            set_shd(c, bg); set_margins(c, 90, 90, 120, 120)
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
    rt.bold = True; rt.font.name = "Calibri"; rt.font.size = Pt(9.5)
    rt.font.color.rgb = RGBColor(43,108,176)
    rb = p.add_run(text)
    rb.font.name = "Calibri"; rb.font.size = Pt(9.5)
    rb.font.color.rgb = RGBColor(45,55,72)
    doc.add_paragraph()

def head(doc, text, lv=1, color=(26,54,93)):
    h = doc.add_heading(level=lv)
    r = h.add_run(text)
    r.font.color.rgb = RGBColor(*color)
    r.font.name = "Calibri"

def title_block(doc, subtitle):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("p096: ATM TRANSACTION FAILURE ANALYSIS")
    r.bold = True; r.font.size = Pt(21); r.font.name = "Calibri"
    r.font.color.rgb = RGBColor(26,54,93)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(subtitle)
    r2.font.size = Pt(13); r2.font.name = "Calibri"
    r2.font.color.rgb = RGBColor(74,85,104)

def meta_block(doc, rows):
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i,(k,v) in enumerate(rows):
        c0,c1 = tbl.rows[i].cells
        set_shd(c0,"F0F4F8"); set_shd(c1,"FFFFFF")
        set_margins(c0,60,60,100,100); set_margins(c1,60,60,100,100)
        r0=c0.paragraphs[0].add_run(k); r0.bold=True
        r0.font.size=Pt(9.5); r0.font.name="Calibri"
        r0.font.color.rgb=RGBColor(26,54,93)
        r1=c1.paragraphs[0].add_run(v)
        r1.font.size=Pt(9.5); r1.font.name="Calibri"
        r1.font.color.rgb=RGBColor(45,55,72)
        c0.width=Inches(2.2); c1.width=Inches(4.3)
    doc.add_paragraph()

def fill_table(table, data):
    for ri, row in enumerate(data, 1):
        for ci, txt in enumerate(row):
            table.rows[ri].cells[ci].paragraphs[0].text = txt

# ══════════════════════════════════════════════════════════
#  STEP 5  –  BRD
# ══════════════════════════════════════════════════════════
def build_brd(out):
    doc = Document()
    for s in doc.sections:
        s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Inches(1.0)

    title_block(doc, "Step 5 — Business Requirements Document (BRD)  v1.0")
    meta_block(doc, [
        ("Project & Code",   "p096 | ATM Transaction Failure Analysis"),
        ("Author / Student", "Lavi Tewari  |  Roll: 2500320140085  |  MCA 2nd Yr DA, ABES"),
        ("Mentor / Reviewer","Mr. Dinesh Sereno  (Academic Project Mentor, ABES)"),
        ("Document Version", "v1.0  —  Baseline (Steps 1–4 integrated)"),
        ("Classification",   "Internal – Academic Submission Only"),
        ("Date Issued",      "November 2026"),
        ("Review Cycle",     "Updated at Steps 6, 8, 11, 13 as RTM progresses"),
    ])

    # ── 1. DOCUMENT CONTROL ──────────────────────────────
    head(doc,"1. Document Control & Change History")
    dc = doc.add_table(rows=4, cols=5)
    fill_table(dc, [
        ("v0.1","12-Nov-2026","Lavi Tewari","Mr. D. Sereno","Initial draft – Steps 1–3 scope"),
        ("v0.9","18-Nov-2026","Lavi Tewari","Mr. D. Sereno","Added stakeholder RACI, risks, glossary"),
        ("v1.0","21-Nov-2026","Lavi Tewari","Mr. D. Sereno","Baselined – all BRs reviewed & approved"),
    ])
    dc.rows[0].cells[0].paragraphs[0].text = "Version"
    dc.rows[0].cells[1].paragraphs[0].text = "Date"
    dc.rows[0].cells[2].paragraphs[0].text = "Author"
    dc.rows[0].cells[3].paragraphs[0].text = "Reviewer"
    dc.rows[0].cells[4].paragraphs[0].text = "Change Summary"
    stylize(dc,[0.7,1.1,1.3,1.3,2.1])
    doc.add_paragraph()

    # ── 2. EXECUTIVE SUMMARY ─────────────────────────────
    head(doc,"2. Executive Summary")
    doc.add_paragraph(
        "This Business Requirements Document formalises the analytical intelligence needs of a retail banking "
        "ATM fleet operations team. Project p096 delivers a Python-based Streamlit web analytics dashboard that "
        "consolidates transaction failure telemetry from a 50 000-record synthetic ATM operational dataset into "
        "actionable insight for fleet managers, NOC engineers, field service leads, and dispute resolution officers. "
        "The dashboard applies structured failure taxonomies, statistical outlier detection, and interactive "
        "Plotly visualisations to allow stakeholders to pinpoint chronic malfunctioning terminals, understand "
        "temporal failure patterns, and quantify the financial value of service disruptions — transitioning "
        "operations from reactive incident response to data-driven preventive management."
    ).style.font.name = "Calibri"

    # ── 3. BUSINESS PROBLEM & OPPORTUNITY ────────────────
    head(doc,"3. Business Problem and Strategic Opportunity")
    doc.add_paragraph(
        "PROBLEM: ATM fleet managers and NOC support teams currently operate without a unified, real-time "
        "diagnostic view of transaction failure events. Terminal logs remain siloed across proprietary switch "
        "controllers and regional data centres, forcing analysts to execute manual SFTP log extractions and "
        "spreadsheet concatenation that can take 24–72 hours to surface a clear failure pattern."
    )
    doc.add_paragraph(
        "OPPORTUNITY: Centralising and structuring 90 days of synthetic ATM transaction data into an interactive "
        "analytical application reduces root-cause identification time by an estimated 70%, enables automated SLA "
        "breach alerts for terminals exceeding the 5% failure-rate threshold, and supports evidence-based "
        "procurement and warranty enforcement decisions worth ₹10–50 lakh per annual hardware refresh cycle."
    )

    # ── 4. OBJECTIVES & KPIs ─────────────────────────────
    head(doc,"4. Business Objectives and Key Performance Indicators (KPIs)")
    kpi = doc.add_table(rows=6, cols=4)
    kpi.rows[0].cells[0].paragraphs[0].text = "Objective"
    kpi.rows[0].cells[1].paragraphs[0].text = "KPI"
    kpi.rows[0].cells[2].paragraphs[0].text = "Target"
    kpi.rows[0].cells[3].paragraphs[0].text = "Measurement Method"
    fill_table(kpi,[
        ("OBJ-01  Fleet health visibility","Fleet-wide failure rate trend refresh time","< 2.0 s page render at 50 000 rows","Streamlit profiler + browser DevTools"),
        ("OBJ-02  Failure taxonomy coverage","% error codes classified (0 unhandled)","100 % mapping accuracy","SQL COUNT of NULL failure_category"),
        ("OBJ-03  Outlier terminal detection","Precision of >5 % SLA breach flag","100 % correct flagging","pytest fixture vs ground-truth labels"),
        ("OBJ-04  Diagnostic UX efficiency","User clicks to isolate worst terminal","≤ 3 clicks from landing page","Usability walkthrough protocol"),
        ("OBJ-05  Test coverage","pytest-cov statement coverage, core modules","≥ 85 % coverage, 0 failures","CI pytest-cov report"),
    ])
    stylize(kpi,[1.4,2.1,1.4,1.6])
    doc.add_paragraph()

    # ── 5. STAKEHOLDERS — RACI ───────────────────────────
    head(doc,"5. Stakeholder Register and RACI Matrix")
    raci = doc.add_table(rows=8, cols=8)
    raci_hdr = ["SH ID","Stakeholder","Role","Interest","Influence",
                "Accountable\n(A)","Consulted\n(C)","Informed\n(I)"]
    for i,h in enumerate(raci_hdr):
        raci.rows[0].cells[i].paragraphs[0].text = h
    raci_data = [
        ("SH-01","ATM Ops Manager","Fleet Ops Lead","High","High","A – BR-01,03,05,08","C – BR-02,07","I – Risk log"),
        ("SH-02","Field Maint. Lead","HW Support / MSP","High","Med","A – BR-03,09,10","C – BR-11,12","I – Deployment"),
        ("SH-03","NOC/Switch Analyst","IT Network Eng.","High","Med","A – BR-04,06,07","C – BR-02","I – Trend alerts"),
        ("SH-04","Dispute Officer","Customer Support","Med","Low","A – BR-05,12","C – BR-08","I – QA reports"),
        ("SH-05","COO","Executive Sponsor","Low","High","A – Budget","C – BR-08,09","I – KPI dashboard"),
        ("SH-06","Academic Mentor","Mr. Dinesh Sereno","High","High","A – R (Responsible-all)","C – All BRs","I – Viva eval"),
        ("Dev","Lavi Tewari","Developer & Analyst","High","High","R – All deliverables","—","—"),
    ]
    fill_table(raci, raci_data)
    stylize(raci,[0.6,1.2,1.1,0.6,0.7,1.2,1.1,1.0])
    doc.add_paragraph()

    callout(doc,
        "RACI KEY — R: Responsible (builds/executes)  |  A: Accountable (signs-off)  "
        "|  C: Consulted (input sought)  |  I: Informed (receives updates).",
        title="LEGEND")

    # ── 6. SCOPE ─────────────────────────────────────────
    head(doc,"6. Project Scope")
    head(doc,"6.1  In-Scope",2)
    for item in [
        "Generation of a 50 000-row synthetic ATM transaction dataset spanning 90 observation days across 5 simulated banks and 200 terminal nodes.",
        "Normalised SQLite relational schema (tables: transactions, atm_terminals, error_codes, maintenance_logs) with FK constraints and B-tree indexes.",
        "Interactive Streamlit multi-page dashboard with sidebar filters (date, bank, terminal, transaction type, failure category).",
        "Six+ Plotly visualisations: fleet failure trend line, failure-category Pareto, hourly failure heatmap, terminal outlier bar chart, latency scatter, On-Us/Off-Us donut.",
        "Statistical modules: failure rate aggregation, outlier detection (μ ± 2σ), latency percentiles (p50/p90/p99), MTBF calculation.",
        "Automated pytest test suite (≥ 85 % coverage) and public deployment on Streamlit Community Cloud.",
        "Full ABES project documentation (Steps 1–18) and viva-ready slide deck.",
    ]:
        p = doc.add_paragraph(style="List Bullet"); p.text = item
        for r in p.runs: r.font.name="Calibri"; r.font.size=Pt(9.5)

    head(doc,"6.2  Out-of-Scope",2)
    for item in [
        "Live ISO 8583 / NDC wire-protocol ATM switch connection — no real banking network integration.",
        "Real PAN, PIN, CVV, or any NPCI-regulated personally identifiable information.",
        "Automated field-engineer dispatch, ServiceNow/Jira real-time bi-directional API integration.",
        "Interbank settlement, clearing, fee calculation, or balance adjustment processing.",
        "Native iOS/Android mobile applications or ATM firmware flashing utilities.",
    ]:
        p = doc.add_paragraph(style="List Bullet"); p.text = item
        for r in p.runs: r.font.name="Calibri"; r.font.size=Pt(9.5)
    doc.add_paragraph()

    # ── 7. BUSINESS REQUIREMENTS ─────────────────────────
    head(doc,"7. Business Requirements Specification")
    doc.add_paragraph(
        "Each requirement is expressed as: 'The business needs to <capability> so that <value>.' "
        "Acceptance Criteria are phrased as verifiable, outcome-based conditions independent of UI technology."
    )
    br_full = [
        ("BR-01","Fleet Failure Rate Trending",
         "The business needs to aggregate ATM transaction failure rates over hourly, daily, and weekly windows so that "
         "fleet managers maintain continuous health visibility.",
         "Closes OQ-01, HMW-01. Fleet managers lack intraday trend visibility until end-of-batch reports.",
         "Must","HMW-01",
         "1. Fleet failure % calculated correctly for any selected date/hour window.\n"
         "2. Metric refresh completes in < 2.0 s on 50 000-row dataset.\n"
         "3. Zero NULL or missing data-points in trend output."),
        ("BR-02","4-Tier Failure Taxonomy Classification",
         "The business needs to map every raw ISO 8583 error code to a 4-tier taxonomy (Hardware / Network-Switch / "
         "Customer-Account / Host-Timeout) so that engineers distinguish machine faults from user errors instantly.",
         "Prevents blame-shifting between HW vendors and IT ops; enables root-cause isolation (OQ-02, HMW-02).",
         "Must","HMW-02",
         "1. 100 % of distinct error codes resolve to a non-NULL failure_category.\n"
         "2. Classification logic unit-tested with ≥ 20 boundary error codes.\n"
         "3. Unknown codes auto-assign to 'Unclassified' and generate a data-quality alert."),
        ("BR-03","Chronic Outlier Terminal Detection",
         "The business needs to auto-detect and rank terminals whose 7-day rolling failure rate exceeds 5.0 % SLA "
         "so that maintenance managers dispatch field engineers to the highest-risk nodes first.",
         "10 % of terminals generate 80 % of customer failure complaints; proactive dispatch saves MSP truck-roll cost (OQ-03, HMW-03).",
         "Must","HMW-03",
         "1. Terminals with failure rate > 5.0 % appear in flagged outlier list with correct rank.\n"
         "2. Outlier algorithm validated against labelled ground-truth fixture (100 % precision).\n"
         "3. Terminals cleared below threshold are automatically removed from alert list."),
        ("BR-04","Temporal Failure Heatmap",
         "The business needs to visualise transaction failure density across a 24-hour × 7-day matrix so that "
         "NOC engineers anticipate congestion peaks and schedule proactive maintenance windows.",
         "Switch operators currently discover Saturday-evening congestion peaks reactively after customer calls (OQ-04, HMW-04).",
         "Should","HMW-04",
         "1. Heatmap renders failure counts across all 24 hours × 7 days without missing cells.\n"
         "2. Peak hour/day cell highlighted with correct failure count label.\n"
         "3. Heatmap updates within 1.5 s on filter change."),
        ("BR-05","Transaction-Type Failure Segmentation",
         "The business needs to segment failure rates across Cash Withdrawal, Balance Inquiry, and PIN Change "
         "operations so that dispute teams prioritise high-impact cash dispense failures.",
         "Charge-back claims for cash traps skew support queue; type-level segmentation enables fair triage (OQ-05, HMW-05).",
         "Must","HMW-05",
         "1. Failure count and rate computed separately for each of the ≥ 3 transaction types.\n"
         "2. Withdrawal identified as highest-volume failure type in test dataset.\n"
         "3. Zero transactions mis-categorised across type boundaries."),
        ("BR-06","Latency-to-Timeout Correlation",
         "The business needs to correlate round-trip network response latency (ms) with timeout failure frequency "
         "so that analysts establish empirical SLA thresholds for switch escalation.",
         "NOC misdiagnoses high-latency periods as terminal hardware offline (OQ-06, HMW-06).",
         "Must","HMW-06",
         "1. Latency percentiles (p50, p90, p99) computed accurately per failure category.\n"
         "2. Pearson/Spearman correlation coefficient between response_time_ms and timeout flag calculated.\n"
         "3. Scatter plot renders without overplotting for 50 000 data points."),
        ("BR-07","On-Us vs Off-Us Network Segregation",
         "The business needs to separate failure metrics by card network type (On-Us same-bank vs Off-Us inter-bank) "
         "so that switch operators isolate external gateway bottlenecks.",
         "Inter-bank NFS switch drops are conflated with domestic core banking issues (OQ-07, HMW-07).",
         "Should","HMW-07",
         "1. On-Us and Off-Us failure rates computed and presented with statistically testable difference.\n"
         "2. Donut/bar chart clearly distinguishes card network volumes.\n"
         "3. Zero transactions assigned to both On-Us and Off-Us simultaneously."),
        ("BR-08","Financial Transaction Value Disruption Analysis",
         "The business needs to quantify total INR transaction value disrupted by hardware faults versus customer "
         "account errors so that executives prioritise remediation investment.",
         "INR value at risk bridges gap between technical KPIs and C-suite financial language (OQ-08, HMW-08).",
         "Must","HMW-08",
         "1. Sum of failed transaction amounts computed correctly per failure category (SQL SUM validated).\n"
         "2. Financial disruption KPI card shows INR values rounded to 2 decimal places.\n"
         "3. Zero approved transactions included in disruption total."),
        ("BR-09","Hardware Vendor Failure Benchmarking",
         "The business needs to rank ATM hardware manufacturers by failure frequency normalised per 1 000 transactions "
         "so that procurement enforces warranty SLAs and informed replacement decisions.",
         "Annual ₹10–50 lakh HW refresh budget lacks objective vendor comparison data (OQ-09, HMW-09).",
         "Should","HMW-09",
         "1. Per-vendor failure rate computed and ranked in descending order.\n"
         "2. Rates normalised per 1 000 transactions to control for fleet size differences.\n"
         "3. Chart annotated with vendor name and model label."),
        ("BR-10","MTBF and Recurrence Frequency Computation",
         "The business needs to compute Mean Time Between Failures (MTBF) and recurrent outage frequency for unstable "
         "terminals so that engineers schedule preventive servicing rather than reactive repairs.",
         "Reactive maintenance cycle doubles MSP on-site costs for intermittent terminals (OQ-10, HMW-10).",
         "Must","HMW-10",
         "1. MTBF computed in running hours for every terminal flagged by BR-03.\n"
         "2. MTBF formula validated: MTBF = (Total observation window − Downtime) / Failure count.\n"
         "3. Terminals sorted by ascending MTBF (most critical first)."),
        ("BR-11","Dynamic Multi-Criteria Filtering",
         "The business needs multi-criteria sidebar filters (date range, bank, terminal, location zone, failure "
         "category) with < 1.5 s re-render so that analysts isolate incidents in ≤ 3 UI interactions.",
         "Usability requirement ensuring all analytic BRs are accessible interactively (HMW-01, HMW-03).",
         "Could","HMW-01, HMW-03",
         "1. All filter combinations produce correct filtered dataset.\n"
         "2. Re-render completes in < 1.5 s for any valid filter permutation.\n"
         "3. Resetting filters to default restores full 50 000-row dataset."),
        ("BR-12","CSV / Excel Export of Filtered Results",
         "The business needs to export filtered transaction subsets and terminal audit summaries as CSV/Excel files "
         "within 2.0 s so that field engineers can carry offline inspection sheets to site.",
         "Field engineers operate in low-connectivity branches; offline data export is a critical workflow aid (HMW-03, HMW-08).",
         "Could","HMW-03, HMW-08",
         "1. Downloaded CSV/Excel contains all columns visible in the filtered table.\n"
         "2. Export completes within 2.0 s for up to 10 000 filtered rows.\n"
         "3. File opens correctly in Microsoft Excel and LibreOffice Calc."),
    ]

    for br in br_full:
        br_id,title_,desc,rat,pri,hmw,ac = br
        head(doc, f"7.{br_full.index(br)+1}  {br_id}: {title_}", 2, (26,54,93))
        
        br_tbl = doc.add_table(rows=6, cols=2)
        labels = ["Description","Rationale","MoSCoW Priority","Source HMW","Acceptance Criteria","Data Needed"]
        # pick data columns needed
        from_br_map = {
            "BR-01":("transaction_id, timestamp, status, terminal_id","transactions, atm_terminals"),
            "BR-02":("error_code, failure_category, failure_reason","error_codes, transactions"),
            "BR-03":("terminal_id, status, timestamp, location","transactions, atm_terminals"),
            "BR-04":("timestamp (hr, dow), failure_category","transactions, error_codes"),
            "BR-05":("transaction_type, status, failure_reason, amount","transactions"),
            "BR-06":("response_time_ms, status, error_code","transactions, error_codes"),
            "BR-07":("card_type, issuing_bank, status","transactions"),
            "BR-08":("amount, failure_category, status","transactions, error_codes"),
            "BR-09":("terminal_id, terminal_model, vendor, failure_category","atm_terminals, error_codes"),
            "BR-10":("terminal_id, timestamp, status, maintenance_id","transactions, maintenance_logs"),
            "BR-11":("timestamp, issuing_bank, terminal_id, location, failure_category","all tables"),
            "BR-12":("all filtered columns","all tables"),
        }
        data_attr, data_src = from_br_map.get(br_id,("—","—"))
        values = [desc, rat, pri, hmw, ac, f"Attributes: {data_attr}\nSource: {data_src}"]
        
        for ri,(lbl,val) in enumerate(zip(labels,values)):
            c0,c1 = br_tbl.rows[ri].cells
            set_shd(c0,"EBF4FF"); set_margins(c0,80,80,120,120)
            set_shd(c1,"FFFFFF"); set_margins(c1,80,80,120,120)
            r0=c0.paragraphs[0].add_run(lbl); r0.bold=True
            r0.font.size=Pt(9.5); r0.font.name="Calibri"
            r0.font.color.rgb=RGBColor(26,54,93)
            c1.paragraphs[0].text = val
            for r in c1.paragraphs[0].runs:
                r.font.size=Pt(9); r.font.name="Calibri"
                r.font.color.rgb=RGBColor(45,55,72)
                if lbl=="MoSCoW Priority":
                    r.bold=True
                    if val=="Must": r.font.color.rgb=RGBColor(197,48,48)
                    elif val=="Should": r.font.color.rgb=RGBColor(146,64,14)
                    else: r.font.color.rgb=RGBColor(44,82,130)
            c0.width=Inches(1.5); c1.width=Inches(5.0)
        doc.add_paragraph()

    # ── 8. BUSINESS RULES ────────────────────────────────
    head(doc,"8. Business Rules")
    rules = [
        ("BRU-01","SLA Breach Threshold","A terminal is classified as an outlier (SLA Breach) when its rolling 7-day failure rate exceeds 5.0%. This threshold is non-negotiable and aligns with the RBI ATM Uptime Directive benchmark."),
        ("BRU-02","Failure Taxonomy Exhaustiveness","Every transaction record must resolve to exactly one of four failure categories: Hardware, Network/Switch, Customer/Account, or Host Timeout. No transaction may carry a NULL failure_category after the classification pipeline runs."),
        ("BRU-03","Synthetic Data Bounds","All transaction amounts (INR) must fall in the realistic range ₹100 to ₹25,000 for Withdrawal, ₹0 for Inquiry. No real PAN, PIN, or card data may appear in any dataset file."),
        ("BRU-04","MTBF Calculation Standard","MTBF = (Total observation window in hours − cumulative downtime hours) ÷ failure count. This formula must be applied consistently and unit-tested against hand-calculated fixtures."),
        ("BRU-05","On-Us / Off-Us Definition","An On-Us transaction is one where the issuing bank equals the acquiring bank. All other issuer–acquirer combinations are Off-Us interbank transactions."),
        ("BRU-06","Latency Threshold Warning","A transaction with response_time_ms > 3,000 ms is automatically classified as a 'High Latency' sub-label within the Network/Switch failure category."),
    ]
    rules_tbl = doc.add_table(rows=len(rules)+1, cols=3)
    rules_tbl.rows[0].cells[0].paragraphs[0].text = "Rule ID"
    rules_tbl.rows[0].cells[1].paragraphs[0].text = "Rule Name"
    rules_tbl.rows[0].cells[2].paragraphs[0].text = "Rule Statement & Rationale"
    for ri,(rid,rn,rs) in enumerate(rules,1):
        rules_tbl.rows[ri].cells[0].paragraphs[0].text=rid
        rules_tbl.rows[ri].cells[1].paragraphs[0].text=rn
        rules_tbl.rows[ri].cells[2].paragraphs[0].text=rs
    stylize(rules_tbl,[0.9,1.5,4.1])
    doc.add_paragraph()

    # ── 9. DATA REQUIREMENTS ─────────────────────────────
    head(doc,"9. Data Requirements")
    dr_tbl = doc.add_table(rows=6, cols=5)
    dr_tbl.rows[0].cells[0].paragraphs[0].text = "Data Entity"
    dr_tbl.rows[0].cells[1].paragraphs[0].text = "Key Attributes"
    dr_tbl.rows[0].cells[2].paragraphs[0].text = "Data Source"
    dr_tbl.rows[0].cells[3].paragraphs[0].text = "Volume"
    dr_tbl.rows[0].cells[4].paragraphs[0].text = "Refresh Strategy"
    dr_data = [
        ("ATM Transactions","transaction_id, terminal_id, timestamp, transaction_type, amount, status, error_code, response_time_ms, card_type, issuing_bank",
         "Synthetic generator (Step 10)","50 000 rows / 90 days","Static seed; re-generated on schema change"),
        ("ATM Terminals","terminal_id, location, city, zone, vendor, terminal_model, install_date, bank_owner",
         "Synthetic generator (Step 10)","200 terminals","Static reference table"),
        ("Error Codes","error_code, description, failure_category, iso_8583_name",
         "ISO 8583 / NPCI NFS standard mapping","~30 distinct codes","Manually curated; version-controlled"),
        ("Maintenance Logs","maintenance_id, terminal_id, event_type, event_datetime, engineer_id, resolution",
         "Synthetic generator (Step 10)","~1 500 events","Static; linked to terminal failure timestamps"),
        ("Failure Classifications","failure_category, sub_category, example_codes",
         "Derived from error_codes table","4 categories, ~12 sub-categories","Auto-derived on data load"),
    ]
    fill_table(dr_tbl, dr_data)
    stylize(dr_tbl,[1.3,2.5,1.3,0.9,1.5])
    doc.add_paragraph()

    # ── 10. ASSUMPTIONS, CONSTRAINTS, DEPENDENCIES ───────
    head(doc,"10. Assumptions, Constraints, and Dependencies")
    head(doc,"10.1  Assumptions",2)
    for a in [
        "The 90-day synthetic dataset faithfully represents real-world multi-bank ATM failure distributions (85–90 % success, 10–15 % failures).",
        "ISO 8583 response code conventions (00=Approval, 51=NSF, 55=Bad PIN, 91=Switch Inoperative, 96=Malfunction) are universally applied.",
        "5.0 % rolling failure rate is accepted as the operational SLA breach trigger aligned with RBI ATM uptime guidance.",
        "Free-tier Streamlit Community Cloud (1 GB RAM, 1 vCPU) is sufficient for the target dataset scale.",
        "SQLite embedded DB file (data/app.db) is adequate for a read-intensive academic analytics workload.",
    ]:
        p=doc.add_paragraph(style="List Bullet"); p.text=a
        for r in p.runs: r.font.name="Calibri"; r.font.size=Pt(9.5)

    head(doc,"10.2  Constraints",2)
    for c in [
        "No live banking API, switch, or CBS integration — synthetic data only.",
        "No real PII (PAN, PIN, Aadhaar); DPDP Act compliance enforced by design.",
        "Technology stack frozen: Python 3.12, Streamlit, SQLite, Plotly, pytest.",
        "Public repository on GitHub; no proprietary credentials or secrets in-repo.",
        "ABES project submission deadline constrains development timeline.",
    ]:
        p=doc.add_paragraph(style="List Bullet"); p.text=c
        for r in p.runs: r.font.name="Calibri"; r.font.size=Pt(9.5)

    head(doc,"10.3  Dependencies",2)
    dep_tbl = doc.add_table(rows=6, cols=3)
    dep_tbl.rows[0].cells[0].paragraphs[0].text = "Dependency"
    dep_tbl.rows[0].cells[1].paragraphs[0].text = "Type"
    dep_tbl.rows[0].cells[2].paragraphs[0].text = "Risk if Unavailable"
    fill_table(dep_tbl,[
        ("Streamlit Community Cloud free tier","Platform","Revert to local `streamlit run` demo with screenshots"),
        ("Python 3.12 runtime (64-bit Windows)","Technology","Freeze at 3.11 — no API differences in scope"),
        ("Plotly ≥ 5.18 for fig.update_layout","Library","Pin to 5.17 and test regression"),
        ("ABES project submission portal","Academic","Offline Word + PDF submission as fallback"),
        ("Step 10 synthetic dataset (CSV seed)","Data","All downstream Steps 11–16 blocked without seed"),
    ])
    stylize(dep_tbl,[2.6,1.2,2.7])
    doc.add_paragraph()

    # ── 11. RISKS ────────────────────────────────────────
    head(doc,"11. Business Risks and Mitigation Strategies")
    risk_tbl = doc.add_table(rows=7, cols=5)
    risk_tbl.rows[0].cells[0].paragraphs[0].text = "Risk ID"
    risk_tbl.rows[0].cells[1].paragraphs[0].text = "Risk Description"
    risk_tbl.rows[0].cells[2].paragraphs[0].text = "Likelihood"
    risk_tbl.rows[0].cells[3].paragraphs[0].text = "Impact"
    risk_tbl.rows[0].cells[4].paragraphs[0].text = "Mitigation Strategy"
    fill_table(risk_tbl,[
        ("R-01","Synthetic data does not reflect realistic failure distributions, invalidating statistical analysis","Low","High","Calibrate generator to industry benchmarks: 10–15 % failure, with 60/20/15/5 % category split"),
        ("R-02","Streamlit Community Cloud free-tier OOM crash on 50 000-row dataset","Medium","High","Apply st.cache_data, paginated SQL queries, and chunked Pandas reads; test on constrained VM"),
        ("R-03","SQLite concurrent read contention if multi-user cloud access","Low","Medium","Enforce read-only connection; use in-memory Pandas for all analysis post-load"),
        ("R-04","Latency column missing or corrupted in generated dataset","Medium","Medium","pytest fixture validates schema on startup; fallback to computed_latency = NULL flag"),
        ("R-05","Academic scope creep inflating project beyond semester deadline","Medium","Medium","RTM status gate: only 'Identified' BRs proceed; 'Could' items deferred if timeline slips"),
        ("R-06","Plagiarism / academic integrity violation from AI-assisted code","Low","Very High","All code hand-reviewed, documented, and viva-explained; AI used for boilerplate only"),
    ])
    stylize(risk_tbl,[0.6,2.3,0.9,0.7,2.0])
    doc.add_paragraph()

    # ── 12. GLOSSARY ─────────────────────────────────────
    head(doc,"12. Glossary of Terms and Acronyms")
    gloss_tbl = doc.add_table(rows=15, cols=2)
    gloss_tbl.rows[0].cells[0].paragraphs[0].text = "Term / Acronym"
    gloss_tbl.rows[0].cells[1].paragraphs[0].text = "Definition"
    glossary = [
        ("ATM","Automated Teller Machine — a self-service banking terminal for cash dispense, balance inquiry, and PIN management."),
        ("MTBF","Mean Time Between Failures — average operational time between successive failure events for a terminal."),
        ("SLA","Service-Level Agreement — contractual uptime/performance targets; breach defined here as > 5 % failure rate."),
        ("On-Us","Transaction where the issuing bank and acquiring ATM operator are the same institution."),
        ("Off-Us","Transaction where the issuing bank and acquiring ATM operator are different; routed via interbank switch."),
        ("NFS","National Financial Switch — NPCI's interbank ATM switching infrastructure in India."),
        ("ISO 8583","International standard defining message formats for card-originated financial transactions."),
        ("p50 / p90 / p99","Latency percentiles: 50th (median), 90th, 99th — standard network SLA measurement markers."),
        ("RTM","Requirements Traceability Matrix — living artefact mapping OQ → HMW → BR → Feature → DB → Test."),
        ("EDA","Exploratory Data Analysis — statistical summarisation and visual inspection of a dataset."),
        ("HMW","How Might We — Design Thinking question frame transforming pain points into opportunity statements."),
        ("NPCI","National Payments Corporation of India — authority governing NFS and UPI payment rails."),
        ("MoSCoW","Must / Should / Could / Won't — agile requirements prioritisation framework."),
        ("CSV","Comma-Separated Values — plain-text tabular data format for portable data export."),
    ]
    fill_table(gloss_tbl, glossary)
    stylize(gloss_tbl,[1.5,5.0])
    doc.add_paragraph()

    # ── 13. SIGN-OFF ─────────────────────────────────────
    head(doc,"13. Document Sign-Off")
    so_tbl = doc.add_table(rows=4, cols=4)
    so_tbl.rows[0].cells[0].paragraphs[0].text = "Role"
    so_tbl.rows[0].cells[1].paragraphs[0].text = "Name"
    so_tbl.rows[0].cells[2].paragraphs[0].text = "Signature / Date"
    so_tbl.rows[0].cells[3].paragraphs[0].text = "Status"
    fill_table(so_tbl,[
        ("Author / Developer","Lavi Tewari  (Roll 2500320140085)","21-Nov-2026","Submitted"),
        ("Project Mentor / Reviewer","Mr. Dinesh Sereno, ABES","_______________","Pending Review"),
        ("Department Head / HOD","ABES Dept. of CS & IT","_______________","Pending Review"),
    ])
    stylize(so_tbl,[1.5,2.0,1.8,1.2])
    callout(doc,
        "This BRD is the authoritative business requirements baseline for project p096. "
        "Any post-sign-off requirement changes must follow the RTM Change-Request process "
        "and are version-controlled in the document header table.",
        title="DOCUMENT BASELINE NOTICE")

    doc.save(out)
    print(f"BRD saved → {out}")


# ══════════════════════════════════════════════════════════
#  STEP 7  –  TRD  (Technical Requirements Document)
# ══════════════════════════════════════════════════════════
def build_trd(out):
    doc = Document()
    for s in doc.sections:
        s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Inches(1.0)

    title_block(doc, "Step 7 — Technical Requirements Document (TRD)  v1.0")
    meta_block(doc, [
        ("Project & Code",   "p096 | ATM Transaction Failure Analysis"),
        ("Author / Student", "Lavi Tewari  |  Roll: 2500320140085  |  MCA 2nd Yr DA, ABES"),
        ("Mentor / Reviewer","Mr. Dinesh Sereno"),
        ("Stack Frozen",     "Python 3.12 · Streamlit · Pandas · NumPy · SciPy · Plotly · SQLite · pytest"),
        ("Document Version", "v1.0  —  Baseline (Steps 1–6 integrated)"),
        ("Date",             "November 2026"),
    ])

    # ── 1. ARCHITECTURE ───────────────────────────────────
    head(doc,"1. System Architecture Overview")
    doc.add_paragraph(
        "The application follows a layered, unidirectional data-flow architecture: "
        "Data Layer → Load & Cache Layer → Transformation Layer → Analysis Layer → Presentation Layer. "
        "No server-side state mutates the SQLite database at runtime; all writes occur offline during "
        "data-generation (Step 10). The Streamlit multi-page runner acts as the application entry point."
    )

    head(doc,"1.1  Architecture Layers (Mermaid Diagram Specification)",2)
    mermaid_spec = (
        "flowchart TD\n"
        "  A[(SQLite\\ndata/app.db)] -->|read-only SQL| B[DataLoader\\nsrc/db.py]\n"
        "  B -->|st.cache_data| C[Raw DataFrames\\n50 000 rows]\n"
        "  C --> D[Transformer\\nsrc/transform.py\\nclean · classify · enrich]\n"
        "  D --> E[Analyser\\nsrc/analysis.py\\naggregations · stats · MTBF]\n"
        "  E --> F{Streamlit Router\\napp.py}\n"
        "  F --> G[pages/01_Fleet_Overview.py]\n"
        "  F --> H[pages/02_Failure_Category.py]\n"
        "  F --> I[pages/03_Terminal_Health.py]\n"
        "  F --> J[pages/04_Temporal_Patterns.py]\n"
        "  F --> K[pages/05_Network_Latency.py]\n"
        "  F --> L[pages/06_Financial_Impact.py]\n"
        "  G & H & I & J & K & L --> M[Plotly Charts]\n"
        "  G & H & I & J & K & L --> N[Tailwind KPI Cards\\nst.components.v1.html]\n"
        "  style A fill:#EBF8FF,stroke:#2B6CB0\n"
        "  style E fill:#FAF5FF,stroke:#6B46C1\n"
        "  style F fill:#F0FFF4,stroke:#276749\n"
    )
    p_mermaid = doc.add_paragraph()
    r_m = p_mermaid.add_run("Mermaid Diagram Source (render at mermaid.live):\n" + mermaid_spec)
    r_m.font.name = "Courier New"; r_m.font.size = Pt(8)
    r_m.font.color.rgb = RGBColor(45,55,72)
    doc.add_paragraph()

    # ── 2. FOLDER STRUCTURE ────────────────────────────────
    head(doc,"2. Repository Folder Structure")
    folder_text = (
        "atm-transaction-analysis/\n"
        "├── app.py                   # Streamlit entry point & sidebar filter engine\n"
        "├── pages/\n"
        "│   ├── 01_Fleet_Overview.py        # OQ-01, BR-01  — trend charts\n"
        "│   ├── 02_Failure_Category.py      # OQ-02, BR-02  — Pareto + taxonomy\n"
        "│   ├── 03_Terminal_Health.py       # OQ-03,09,10, BR-03,09,10 — outlier table\n"
        "│   ├── 04_Temporal_Patterns.py     # OQ-04, BR-04  — 24×7 heatmap\n"
        "│   ├── 05_Network_Latency.py       # OQ-06,07, BR-06,07 — scatter & donut\n"
        "│   └── 06_Financial_Impact.py      # OQ-05,08, BR-05,08 — value KPIs\n"
        "├── src/\n"
        "│   ├── __init__.py\n"
        "│   ├── db.py                # SQLite connection, cached loaders\n"
        "│   ├── transform.py         # Cleaning, taxonomy mapping, feature engineering\n"
        "│   ├── analysis.py          # Aggregations, outlier detection, MTBF, stats\n"
        "│   └── charts.py            # Plotly figure factory functions\n"
        "├── data/\n"
        "│   ├── app.db               # SQLite database (generated by scripts/seed.py)\n"
        "│   └── seed_config.json     # Synthetic data generator parameters\n"
        "├── scripts/\n"
        "│   └── seed.py              # Synthetic dataset generator (Step 10)\n"
        "├── tests/\n"
        "│   ├── conftest.py          # pytest fixtures (in-memory SQLite, sample df)\n"
        "│   ├── test_db.py           # DB loader tests\n"
        "│   ├── test_transform.py    # Classification & cleaning tests\n"
        "│   ├── test_analysis.py     # Aggregation, outlier, MTBF tests\n"
        "│   └── test_ui.py           # Streamlit AppTest smoke tests\n"
        "├── docs/                    # Steps 1–18 Word/Excel documents\n"
        "├── requirements.txt         # Pinned dependency manifest\n"
        "├── .streamlit/\n"
        "│   └── config.toml          # Streamlit theme + server settings\n"
        "└── README.md"
    )
    p_folder = doc.add_paragraph()
    r_f = p_folder.add_run(folder_text)
    r_f.font.name = "Courier New"; r_f.font.size = Pt(8)
    r_f.font.color.rgb = RGBColor(45,55,72)
    doc.add_paragraph()

    # ── 3. FUNCTIONAL REQUIREMENTS (FR) ──────────────────
    head(doc,"3. Functional Requirements (FR-01 to FR-16)")
    doc.add_paragraph("Each FR traces directly to one or more Business Requirements (BR) and identifies the implementing Python module:")
    
    fr_tbl = doc.add_table(rows=17, cols=5)
    fr_hdr = ["FR ID","Functional Requirement","Linked BR","Implementing Module","Output Artifact"]
    for i,h in enumerate(fr_hdr):
        fr_tbl.rows[0].cells[i].paragraphs[0].text = h
    fr_data = [
        ("FR-01","Load all 4 SQLite tables into Pandas DataFrames with st.cache_data TTL","BR-01..12","src/db.py · load_transactions() · load_terminals()","Cached DataFrames"),
        ("FR-02","Validate schema: assert required columns exist and no NULL transaction_id","BR-02","src/db.py · validate_schema()","ValueError on mismatch"),
        ("FR-03","Map every error_code to failure_category using error_codes lookup join","BR-02","src/transform.py · classify_failures()","transactions df + failure_category col"),
        ("FR-04","Compute hourly, daily, weekly failure rates (success/total per window)","BR-01","src/analysis.py · compute_failure_rate()","Aggregated Series/df"),
        ("FR-05","Flag outlier terminals: rolling 7-day failure rate > 5.0 %","BR-03","src/analysis.py · detect_outliers()","outlier_terminals df + bool flag col"),
        ("FR-06","Generate 24-hour × 7-day failure density heatmap data matrix","BR-04","src/analysis.py · temporal_heatmap()","Pivot table (hour × dow)"),
        ("FR-07","Compute failure breakdown by transaction_type (Withdrawal/Inquiry/PIN)","BR-05","src/analysis.py · failure_by_type()","Grouped df with count + rate"),
        ("FR-08","Compute latency percentiles (p50/p90/p99) per failure_category","BR-06","src/analysis.py · latency_percentiles()","Dict of percentile values"),
        ("FR-09","Compute Pearson correlation: response_time_ms vs timeout_flag","BR-06","src/analysis.py · latency_correlation()","float r-value + p-value"),
        ("FR-10","Segregate failure counts and rates by card_type (On-Us / Off-Us)","BR-07","src/analysis.py · onus_vs_offus()","Grouped df"),
        ("FR-11","Sum INR transaction value of failed records grouped by failure_category","BR-08","src/analysis.py · financial_disruption()","Aggregated df with sum_amount"),
        ("FR-12","Rank hardware vendors by failure rate normalised per 1 000 transactions","BR-09","src/analysis.py · vendor_benchmark()","Ranked df"),
        ("FR-13","Compute MTBF per terminal using timestamped failure event deltas","BR-10","src/analysis.py · compute_mtbf()","df with mtbf_hours col"),
        ("FR-14","Render all Plotly charts via centralized chart factory","BR-01..10","src/charts.py · * factory functions","Plotly Figure objects"),
        ("FR-15","Apply Streamlit sidebar filters; pass filtered df to all analysis functions","BR-11","app.py · sidebar_filter_engine()","Filtered DataFrame"),
        ("FR-16","Provide CSV/Excel download of filtered DataFrame via st.download_button","BR-12","pages/*.py · export_section()","Binary file download stream"),
    ]
    fill_table(fr_tbl, fr_data)
    stylize(fr_tbl,[0.7,2.5,0.9,1.5,1.0])
    doc.add_paragraph()

    # ── 4. DATA LAYER ─────────────────────────────────────
    head(doc,"4. Data Layer Specification")
    head(doc,"4.1  SQLite Schema Summary",2)
    doc.add_paragraph(
        "Database file: data/app.db  |  Engine: SQLite 3.45+ (Python built-in sqlite3)  "
        "|  Mode: READ-ONLY at runtime (uri=True, ?mode=ro)  |  Connection pool: none — single thread Streamlit."
    )
    head(doc,"4.2  Caching Strategy",2)
    doc.add_paragraph(
        "@st.cache_data(ttl=3600) wraps all db.py loader functions. "
        "Full DataFrames are cached in Streamlit session state for 1 hour. "
        "Filter operations occur in-memory on cached copies — zero additional DB round-trips."
    )
    head(doc,"4.3  Validation & Data Quality",2)
    for rule in [
        "Schema assertion: assert set(REQUIRED_COLS).issubset(df.columns) — raises ValueError logged to Streamlit error widget.",
        "NULL gate: assert df['transaction_id'].notna().all() — no orphan transaction rows allowed.",
        "Range check: assert df['amount'].between(0, 25000).all() — data quality warning if violated.",
        "FK integrity: every terminal_id in transactions must exist in atm_terminals (LEFT JOIN NULL check).",
        "Timestamp format: enforce pd.to_datetime() coercion; invalid rows written to data/quality_log.csv.",
    ]:
        p=doc.add_paragraph(style="List Bullet"); p.text=rule
        for r in p.runs: r.font.name="Calibri"; r.font.size=Pt(9.5)
    doc.add_paragraph()

    # ── 5. STATISTICAL METHODS ────────────────────────────
    head(doc,"5. Statistical Method per Operational Question")
    stat_tbl = doc.add_table(rows=11, cols=4)
    stat_tbl.rows[0].cells[0].paragraphs[0].text = "OQ"
    stat_tbl.rows[0].cells[1].paragraphs[0].text = "Statistical / Analytical Method"
    stat_tbl.rows[0].cells[2].paragraphs[0].text = "Python Implementation"
    stat_tbl.rows[0].cells[3].paragraphs[0].text = "Output"
    stat_data = [
        ("OQ-01","Time-series groupby aggregation — failure rate = failed / total per temporal window",
         "df.groupby(pd.Grouper(freq='H'))['status'].apply(lambda x: (x=='Fail').sum()/len(x))","Line chart + KPI tile"),
        ("OQ-02","Frequency distribution + Pareto 80/20 cumulative share by failure_category",
         "df.groupby('failure_category').size().sort_values(ascending=False); cumsum / total","Pareto bar+line chart"),
        ("OQ-03","Outlier detection: rolling 7-day failure rate > μ + 2σ AND > 5.0 % threshold",
         "df.groupby('terminal_id').rolling('7D').apply(); scipy.stats.zscore()","Flagged terminal table"),
        ("OQ-04","Two-dimensional pivot — failure counts × (hour of day, day of week)",
         "df.pivot_table(index='hour', columns='dow', values='status', aggfunc='count')","Plotly Heatmap"),
        ("OQ-05","Grouped proportions — failure rate per transaction_type",
         "df.groupby('transaction_type')['status'].value_counts(normalize=True)","Stacked/grouped bar"),
        ("OQ-06","Latency percentile (np.percentile) + Pearson / Spearman correlation with timeout flag",
         "np.percentile(df['response_time_ms'], [50,90,99]); scipy.stats.pearsonr()","Scatter + r-value KPI"),
        ("OQ-07","Proportion test: compare On-Us vs Off-Us failure rates; two-proportion Z-test",
         "statsmodels.stats.proportion.proportions_ztest([fail_onus, fail_offus], [n_onus, n_offus])","Donut + Z-stat"),
        ("OQ-08","Group sum of transaction amounts by failure_category (failed records only)",
         "df[df['status']=='Fail'].groupby('failure_category')['amount'].sum()","KPI cards (INR)"),
        ("OQ-09","Normalised failure rate per 1 000 transactions per vendor × model",
         "df.groupby('vendor').apply(lambda g: g['status'].eq('Fail').sum()/len(g)*1000)","Ranked bar chart"),
        ("OQ-10","MTBF = (window_hours − downtime_hours) / failure_count per terminal",
         "df.sort_values('timestamp').groupby('terminal_id').apply(mtbf_lambda)","MTBF table + bar"),
    ]
    fill_table(stat_tbl, stat_data)
    stylize(stat_tbl,[0.7,2.0,2.4,1.4])
    doc.add_paragraph()

    # ── 6. UI NOTES ───────────────────────────────────────
    head(doc,"6. UI Framework and Design Notes")
    head(doc,"6.1  Streamlit Multi-Page Structure",2)
    doc.add_paragraph(
        "app.py serves as the shell: Streamlit sidebar hosts all shared filters "
        "(st.date_input, st.multiselect, st.slider). Filtered DataFrame is stored in st.session_state['filtered_df'] "
        "and accessed by all pages via import app.get_filtered_df()."
    )
    head(doc,"6.2  Tailwind CSS Integration",2)
    doc.add_paragraph(
        "Tailwind CDN is injected via st.components.v1.html() for KPI card tiles and the top navigation header. "
        "Pattern: render HTML string with Tailwind classes (flex, shadow-md, rounded-xl, text-indigo-700) inside "
        "st.components.v1.html(html_string, height=120). Full Streamlit theme follows dark-mode-compatible colour palette."
    )
    head(doc,"6.3  Plotly Theming Standard",2)
    doc.add_paragraph(
        "All charts use a shared layout template (src/charts.py · ATM_THEME dict): "
        "plot_bgcolor='#F7FAFC', paper_bgcolor='#FFFFFF', font_family='Inter, Calibri', "
        "colorway=['#2B6CB0','#E53E3E','#D69E2E','#38A169','#805AD5','#DD6B20']. "
        "Plotly config: {'displayModeBar': True, 'toImageButtonOptions': {'format': 'svg'}}."
    )
    head(doc,"6.4  Responsive Layout",2)
    doc.add_paragraph(
        "Three-column KPI row: st.columns([1,1,1]). Chart container: st.columns([2,1]) for chart + insight panel. "
        "st.set_page_config(layout='wide') enforced. Mobile breakpoint handled by Tailwind sm: prefix classes."
    )
    doc.add_paragraph()

    # ── 7. NFRs ───────────────────────────────────────────
    head(doc,"7. Non-Functional Requirements (Performance, Security, Error Handling)")
    nfr_tbl = doc.add_table(rows=9, cols=4)
    nfr_tbl.rows[0].cells[0].paragraphs[0].text = "NFR ID"
    nfr_tbl.rows[0].cells[1].paragraphs[0].text = "Category"
    nfr_tbl.rows[0].cells[2].paragraphs[0].text = "Requirement"
    nfr_tbl.rows[0].cells[3].paragraphs[0].text = "Verification"
    nfr_data = [
        ("NFR-01","Performance","Page render time < 2.0 s on 50 000-row dataset with @st.cache_data","Streamlit profiler timer + browser DevTools"),
        ("NFR-02","Performance","Filter re-render time < 1.5 s for any valid sidebar selection combination","Stopwatch test across 10 filter permutations"),
        ("NFR-03","Reliability","Application must not throw unhandled exceptions on any valid filter combination","pytest AppTest + manual smoke test for all 6 pages"),
        ("NFR-04","Security","SQLite opened in read-only URI mode: sqlite3.connect('file:data/app.db?mode=ro', uri=True)","Code review + unit test asserting no INSERT/UPDATE executed"),
        ("NFR-05","Security","No PAN, PIN, CVV, Aadhaar, or real user data in any file committed to repo","Git pre-commit hook + manual audit of seed.py output"),
        ("NFR-06","Maintainability","All src/ modules covered ≥ 85 % by pytest-cov; zero star imports","pytest-cov HTML report in CI"),
        ("NFR-07","Portability","requirements.txt with pinned versions; app runs on Python 3.11+ (Linux/Windows/macOS)","Test on fresh venv on each OS before submission"),
        ("NFR-08","Availability","App publicly accessible on Streamlit Community Cloud for ≥ 30 days post-deployment","URL reachability check on viva day"),
    ]
    fill_table(nfr_tbl, nfr_data)
    stylize(nfr_tbl,[0.7,1.1,2.9,1.8])
    doc.add_paragraph()

    # ── 8. REQUIREMENTS.TXT ───────────────────────────────
    head(doc,"8. Pinned requirements.txt")
    reqs = (
        "# p096: ATM Transaction Failure Analysis\n"
        "# Python 3.12 | Pinned for Streamlit Community Cloud deployment\n"
        "streamlit==1.37.0\n"
        "pandas==2.2.2\n"
        "numpy==1.26.4\n"
        "plotly==5.22.0\n"
        "scipy==1.13.1\n"
        "statsmodels==0.14.2\n"
        "openpyxl==3.1.5\n"
        "xlsxwriter==3.2.0\n"
        "python-docx==1.2.0\n"
        "pytest==8.2.2\n"
        "pytest-cov==5.0.0\n"
        "# Note: sqlite3 is part of Python standard library — no pip install required\n"
    )
    p_req = doc.add_paragraph()
    r_req = p_req.add_run(reqs)
    r_req.font.name="Courier New"; r_req.font.size=Pt(8.5)
    r_req.font.color.rgb=RGBColor(45,55,72)
    doc.add_paragraph()

    # ── 9. DEPLOYMENT ─────────────────────────────────────
    head(doc,"9. Deployment Requirements")
    dep_steps = [
        ("DEP-01","Git Repository Setup","Create public GitHub repo: github.com/<user>/atm-transaction-analysis. Push all code, docs/, and data/app.db. Add .gitignore excluding __pycache__, .venv, *.log."),
        ("DEP-02","Streamlit Secrets","No secrets required — SQLite is file-based. Confirm .streamlit/secrets.toml is NOT present in repo."),
        ("DEP-03","Community Cloud Config","Link GitHub repo to Streamlit Community Cloud. Set Main file path: app.py. Python version: 3.12. Click Deploy."),
        ("DEP-04","Smoke Test Post-Deploy","Verify all 6 pages load without error. Confirm filter interactions work. Test CSV download. Validate MTBF and latency chart renders correctly."),
        ("DEP-05","Submission URL","Share public Streamlit app URL in ABES project submission form alongside GitHub repo link."),
    ]
    dep_tbl = doc.add_table(rows=len(dep_steps)+1, cols=3)
    dep_tbl.rows[0].cells[0].paragraphs[0].text = "Step ID"
    dep_tbl.rows[0].cells[1].paragraphs[0].text = "Deployment Activity"
    dep_tbl.rows[0].cells[2].paragraphs[0].text = "Detail"
    fill_table(dep_tbl, dep_steps)
    stylize(dep_tbl,[0.8,1.6,4.1])
    doc.add_paragraph()

    # ── 10. TECHNICAL RISKS ────────────────────────────────
    head(doc,"10. Technical Risks and Mitigation")
    tr_tbl = doc.add_table(rows=7, cols=4)
    tr_tbl.rows[0].cells[0].paragraphs[0].text = "TR ID"
    tr_tbl.rows[0].cells[1].paragraphs[0].text = "Technical Risk"
    tr_tbl.rows[0].cells[2].paragraphs[0].text = "Likelihood / Impact"
    tr_tbl.rows[0].cells[3].paragraphs[0].text = "Mitigation"
    tr_data = [
        ("TR-01","Streamlit AppTest API changes between 1.37 and next minor release","Low / Med","Pin streamlit==1.37.0 strictly; test before upgrade"),
        ("TR-02","Plotly version incompatibility with Pandas 2.x DataFrame API","Low / High","Lock plotly==5.22.0; run cross-version matrix in pytest"),
        ("TR-03","SQLite WAL journal lock during concurrent Streamlit reruns","Med / Med","Use read-only URI mode + single-thread Streamlit config"),
        ("TR-04","Memory spike on 50 000-row full join (terminals + transactions)","Med / High","Pre-join in SQL; expose only required columns to Pandas"),
        ("TR-05","Tailwind CDN blocked on Streamlit Community Cloud CSP headers","Low / Med","Fallback: inline critical CSS styles in st.markdown unsafe_allow_html"),
        ("TR-06","pytest-cov misses Streamlit page files (pages/*.py)","Med / Low","Configure coverage.ini: source=['src','pages']; add AppTest smoke tests"),
    ]
    fill_table(tr_tbl, tr_data)
    stylize(tr_tbl,[0.7,2.5,1.2,2.1])
    doc.add_paragraph()

    callout(doc,
        "RTM UPDATE NOTICE (Step 7): FR-01 through FR-16 are new functional requirement IDs created at this step. "
        "The RTM (04_Traceability_Matrix.xlsx) must be updated at Step 6 (PRD) with Feature column, and again here "
        "with FR IDs in the Feature column. This will be completed in the Step 6 PRD session.",
        title="RTM UPDATE REQUIRED")

    doc.save(out)
    print(f"TRD saved → {out}")


# ── MAIN ─────────────────────────────────────────────────
if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_brd("docs/05_BRD.docx")
    build_trd("docs/07_TRD.docx")
