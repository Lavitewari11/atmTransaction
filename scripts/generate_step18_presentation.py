import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Color Palette ──
RED = RGBColor(0xA5, 0x1C, 0x30)
BLACK = RGBColor(0x15, 0x15, 0x14)
DARK = RGBColor(0x3A, 0x3A, 0x38)
GREY = RGBColor(0x5E, 0x5E, 0x5B)
CARD_BG = RGBColor(0xF4, 0xF1, 0xF1)
PINK = RGBColor(0xD9, 0xB5, 0xBC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER_GREY = RGBColor(0xBF, 0xBF, 0xBF)

def create_presentation(output_path):
    prs = Presentation()
    prs.slide_width = Inches(10.0)
    prs.slide_height = Inches(5.625)
    blank_layout = prs.slide_layouts[6]
    
    # Helper to add standard footer
    def add_footer(slide, slide_num):
        # Footer text
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = "HCLTech Career Shaper ABES Engineering College | p096 - ATM Transaction Failure Analysis"
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.color.rgb = GREY
        
        # Slide number
        tb_num = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        r_num = p_num.add_run()
        r_num.text = str(slide_num)
        r_num.font.name = "Calibri"
        r_num.font.size = Pt(9)
        r_num.font.color.rgb = RGBColor(0x8A, 0x8A, 0x86)

    def add_content_title(slide, title_text):
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = title_text
        r.font.name = "Cambria"
        r.font.size = Pt(28)
        r.font.bold = True
        r.font.color.rgb = BLACK

    def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text):
        # Card rounded rectangle
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = CARD_BG
        shape.line.color.rgb = CARD_BG
        
        # Badge circle
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.2), Inches(y + 0.2), Inches(0.42), Inches(0.42))
        badge.fill.solid()
        badge.fill.fore_color.rgb = RED
        badge.line.color.rgb = RED
        btf = badge.text_frame
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = str(badge_text)
        br.font.name = "Calibri"
        br.font.size = Pt(12)
        br.font.bold = True
        br.font.color.rgb = WHITE
        
        # Head text
        tb_head = slide.shapes.add_textbox(Inches(x + 0.75), Inches(y + 0.17), Inches(w - 0.9), Inches(0.45))
        htf = tb_head.text_frame
        htf.word_wrap = True
        hp = htf.paragraphs[0]
        hp.alignment = PP_ALIGN.LEFT
        hr = hp.add_run()
        hr.text = head_text
        hr.font.name = "Calibri"
        hr.font.size = Pt(14)
        hr.font.bold = True
        hr.font.color.rgb = BLACK
        
        # Body text
        tb_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
        btf2 = tb_body.text_frame
        btf2.word_wrap = True
        bp2 = btf2.paragraphs[0]
        bp2.alignment = PP_ALIGN.LEFT
        br2 = bp2.add_run()
        br2.text = body_text
        br2.font.name = "Calibri"
        br2.font.size = Pt(12)
        br2.font.color.rgb = DARK

    def add_plain_card(slide, x, y, w, h, head_text, body_text):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = CARD_BG
        shape.line.color.rgb = CARD_BG
        
        tb_head = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.17), Inches(w - 0.35), Inches(0.45))
        htf = tb_head.text_frame
        htf.word_wrap = True
        hp = htf.paragraphs[0]
        hr = hp.add_run()
        hr.text = head_text
        hr.font.name = "Calibri"
        hr.font.size = Pt(14)
        hr.font.bold = True
        hr.font.color.rgb = BLACK
        
        tb_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
        btf = tb_body.text_frame
        btf.word_wrap = True
        bp = btf.paragraphs[0]
        br = bp.add_run()
        br.text = body_text
        br.font.name = "Calibri"
        br.font.size = Pt(12)
        br.font.color.rgb = DARK

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 1: DARK (Title Slide)
    # ══════════════════════════════════════════════════════════════════════════
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.background
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = BLACK
    
    # Header logo & programme
    tb_logo = s1.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    p = tb_logo.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "[ABES logo]"
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.font.italic = True
    r.font.color.rgb = PINK
    
    tb_prog = s1.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    p = tb_prog.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.color.rgb = PINK
    
    # Title
    tb_title = s1.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    p = tb_title.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "p096 - ATM Transaction Failure Analysis"
    r.font.name = "Cambria"
    r.font.size = Pt(36)
    r.font.bold = True
    r.font.color.rgb = WHITE
    
    # Subtitle / Details
    tb_det = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_det = tb_det.text_frame
    tf_det.word_wrap = True
    p1 = tf_det.paragraphs[0]
    p1.space_before = Pt(4)
    r = p1.add_run()
    r.text = "Lavi Tewari | Roll No 2500320140085 / MCA 2nd Year Data Analyst"
    r.font.name = "Calibri"
    r.font.size = Pt(15)
    r.font.color.rgb = CARD_BG
    
    p2 = tf_det.add_paragraph()
    p2.space_before = Pt(4)
    r = p2.add_run()
    r.text = "Mentor: Mr. Dinesh Sereno | November 2026"
    r.font.name = "Calibri"
    r.font.size = Pt(15)
    r.font.color.rgb = PINK
    
    s1.notes_slide.notes_text_frame.text = (
        "Good morning everyone. I am Lavi Tewari, MCA 2nd Year Data Analytics student at ABES. "
        "Today I present project p096: ATM Transaction Failure Analysis, developed under mentor Mr. Dinesh Sereno. "
        "This project provides actionable operational intelligence across 50,000 multi-bank transaction events to isolate "
        "terminal downtime and SLA breach root causes."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 2: Problem and How Might We
    # ══════════════════════════════════════════════════════════════════════════
    s2 = prs.slides.add_slide(blank_layout)
    add_content_title(s2, "Problem and How Might We")
    add_footer(s2, 2)
    
    add_badge_card(
        s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem Statement",
        "ATM terminal downtime and recurring transaction failures (technical timeouts, cash dispenser jams, switch link outages) "
        "reduce transaction success rates below acceptable banking SLAs (< 95%), erode customer trust, and inflate maintenance costs "
        "due to the absence of centralized diagnostic visibility across banking terminals."
    )
    add_badge_card(
        s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Baseline Visibility",
        "How might we give fleet operations immediate visibility into real-time failure trends across hourly, daily, and weekly cycles?"
    )
    add_badge_card(
        s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Root-Cause Taxonomy",
        "How might we systematically categorize transaction declines into clear banking root causes (Hardware, Network, Customer, Host Timeout)?"
    )
    s2.notes_slide.notes_text_frame.text = (
        "Here we define our core problem and design-thinking opportunity statements. ATM fleet managers operate blind "
        "to intraday failure surges. By framing HMW-01 and HMW-02, we address both operational trend tracking and "
        "objective 4-tier failure classification."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 3: Stakeholders and users
    # ══════════════════════════════════════════════════════════════════════════
    s3 = prs.slides.add_slide(blank_layout)
    add_content_title(s3, "Stakeholders and users")
    add_footer(s3, 3)
    
    add_badge_card(
        s3, 0.5, 1.35, 2.85, 3.5, "1", "ATM Ops Manager",
        "Primary User (SH-01)\n\n"
        "Monitors fleet availability KPIs, identifies SLA breach hotspots, and tracks cash replenishment efficiency.\n\n"
        "Influence: High"
    )
    add_badge_card(
        s3, 3.55, 1.35, 2.85, 3.5, "2", "Field Maintenance Lead",
        "Secondary User (SH-02)\n\n"
        "Examines localized hardware error codes, prioritizes truck rolls, and schedules preventive servicing.\n\n"
        "Influence: Medium"
    )
    add_badge_card(
        s3, 6.6, 1.35, 2.85, 3.5, "3", "Academic Mentor",
        "Evaluator (SH-06)\n\n"
        "Evaluates data engineering rigor, statistical distributions, software hygiene, and ABES curriculum criteria.\n\n"
        "Influence: High"
    )
    s3.notes_slide.notes_text_frame.text = (
        "Our stakeholder ecosystem addresses operational, hardware, and academic needs. The Operations Manager needs real-time "
        "SLA monitoring, the Field Maintenance Lead requires targeted machine diagnostic data for dispatch, and our academic mentor "
        "evaluates methodology and code hygiene."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 4: Requirements and traceability
    # ══════════════════════════════════════════════════════════════════════════
    s4 = prs.slides.add_slide(blank_layout)
    add_content_title(s4, "Requirements and traceability")
    add_footer(s4, 4)
    
    # Table at (.5,1.3), col widths 1, 3.4, 1, 1.2, 1.2, 1.2
    tbl_shape = s4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(1.0)
    tbl.columns[1].width = Inches(3.4)
    tbl.columns[2].width = Inches(1.0)
    tbl.columns[3].width = Inches(1.2)
    tbl.columns[4].width = Inches(1.2)
    tbl.columns[5].width = Inches(1.2)
    
    headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = WHITE
        
    rtm_sample = [
        ("HMW-01", "Fleet failure rate trending over time", "Must", "US-01, 04", "TC-F-01, 04", "UI-01"),
        ("HMW-02", "4-tier banking failure taxonomy", "Must", "US-02, 03", "TC-F-02, 07", "UI-02"),
        ("HMW-03", "Outlier terminal detection (>5% SLA)", "Must", "US-10", "TC-F-10", "UI-03"),
        ("HMW-04", "24x7 temporal failure density heatmap", "Should", "US-08", "TC-F-08", "UI-04")
    ]
    for r_idx, r_data in enumerate(rtm_sample, start=1):
        for c_idx, val in enumerate(r_data):
            cell = tbl.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx in (0, 2, 3, 4, 5) else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.font.color.rgb = BLACK
            
    tb_note = s4.shapes.add_textbox(Inches(0.5), Inches(4.3), Inches(9.0), Inches(0.35))
    p = tb_note.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "Full matrix: Project Document, section 4 (12 requirements with 100% test coverage)"
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.italic = True
    r.font.color.rgb = GREY
    
    s4.notes_slide.notes_text_frame.text = (
        "This slide demonstrates strict requirement engineering. Every HMW flows directly to a testable Business Requirement, "
        "User Story, automated test case, and UI screen. Zero unmapped requirements enter our build plan."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 5: Data and preparation
    # ══════════════════════════════════════════════════════════════════════════
    s5 = prs.slides.add_slide(blank_layout)
    add_content_title(s5, "Data and preparation")
    add_footer(s5, 5)
    
    stats_data = [
        (0.5, "50,003", "Total Transactions"),
        (2.8, "35", "Analytical Columns"),
        (5.1, "90 Days", "Observation Window"),
        (7.4, "1,018", "Data Issues Quarantined")
    ]
    for x, val, lbl in stats_data:
        tb_val = s5.shapes.add_textbox(Inches(x), Inches(1.3), Inches(2.1), Inches(0.8))
        p = tb_val.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = val
        r.font.name = "Cambria"
        r.font.size = Pt(28)
        r.font.bold = True
        r.font.color.rgb = RED
        
        tb_lbl = s5.shapes.add_textbox(Inches(x), Inches(2.1), Inches(2.1), Inches(0.35))
        p = tb_lbl.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = lbl
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.color.rgb = GREY
        
    add_plain_card(
        s5, 0.5, 2.7, 4.3, 2.25, "Data Source & Representation",
        "Synthetic retail banking dataset modeling 200 ATM terminals across 20 Indian cities and 5 member banks (SBI, HDFC, ICICI, PNB, BOI). "
        "Each record captures an individual transaction event with ISO 8583 response codes, latency in ms, and card scheme routing."
    )
    add_plain_card(
        s5, 5.2, 2.7, 4.3, 2.25, "Cleaning & Data Hygiene",
        "1. Quarantined 1,000 missing latency values to prevent metric skew.\n"
        "2. De-duplicated 3 duplicate transaction IDs during ingestion.\n"
        "3. Filtered 15 extreme amount outliers (> ₹50,000) into dedicated audit log."
    )
    s5.notes_slide.notes_text_frame.text = (
        "Our synthetic dataset reflects realistic operational parameters over a 90-day window. We generated 50,003 records "
        "and intentionally embedded realistic anomalies to demonstrate robust data validation and cleaning pipelines."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 6: Solution and technology
    # ══════════════════════════════════════════════════════════════════════════
    s6 = prs.slides.add_slide(blank_layout)
    add_content_title(s6, "Solution and technology")
    add_footer(s6, 6)
    
    steps = [
        (0.5, "Raw Data\n(data/app.db)", False),
        (2.35, "Clean & KPIs\n(Python, Pandas)", False),
        (4.2, "Store & Cache\n(SQLite / Session)", False),
        (6.05, "Dashboard\n(Streamlit)", True),
        (7.9, "Cloud Deploy\n(Community Cloud)", False)
    ]
    for x, txt, is_red in steps:
        box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.5), Inches(1.6), Inches(1.2))
        box.fill.solid()
        box.fill.fore_color.rgb = RED if is_red else CARD_BG
        box.line.color.rgb = RED if is_red else CARD_BG
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = txt
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = WHITE if is_red else BLACK
        
    arrows_x = [2.1, 3.95, 5.8, 7.65]
    for ax in arrows_x:
        tb = s6.shapes.add_textbox(Inches(ax), Inches(1.5), Inches(0.25), Inches(1.2))
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ">"
        r.font.name = "Calibri"
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = RED
        
    tb_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    p = tb_tools.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r_lbl = p.add_run()
    r_lbl.text = "Tools: "
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = DARK
    r_tools = p.add_run()
    r_tools.text = "Python 3.11 · Streamlit 1.37 · Pandas · NumPy · Plotly · SQLite3 · pytest · Tailwind CSS"
    r_tools.font.color.rgb = DARK
    
    tb_cmd = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    p = tb_cmd.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r_cmd = p.add_run()
    r_cmd.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    r_cmd.font.italic = True
    r_cmd.font.color.rgb = GREY
    
    s6.notes_slide.notes_text_frame.text = (
        "Our solution pipeline is clean, layered, and completely zero-cost. We process raw transactional records, cache them "
        "in-memory, render interactive Plotly analytics through Streamlit with custom Tailwind styling, and deploy on Community Cloud."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 7: Dashboard
    # ══════════════════════════════════════════════════════════════════════════
    s7 = prs.slides.add_slide(blank_layout)
    add_content_title(s7, "Dashboard")
    add_footer(s7, 7)
    
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid()
    frame.fill.fore_color.rgb = CARD_BG
    frame.line.color.rgb = BORDER_GREY
    frame.line.dash_style = 2
    
    tf = frame.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "[Paste a dashboard screenshot here]"
    r.font.name = "Calibri"
    r.font.size = Pt(14)
    r.font.italic = True
    r.font.color.rgb = GREY
    
    nav_screens = [
        (1.3, "1", "UI-01", "Fleet Overview: Real-time KPIs & Hourly Trend"),
        (2.2, "2", "UI-02", "Failure Taxonomy: Pareto 80/20 & Type Split"),
        (3.1, "3", "UI-03", "Terminal Health: Outlier Table (>5%) & MTBF"),
        (4.0, "4", "UI-04", "Temporal Patterns: 24x7 Density Heatmap")
    ]
    for y, num, tag, desc in nav_screens:
        badge = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.7), Inches(y), Inches(0.42), Inches(0.42))
        badge.fill.solid()
        badge.fill.fore_color.rgb = RED
        badge.line.color.rgb = RED
        bp = badge.text_frame.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = num
        br.font.size = Pt(12)
        br.font.bold = True
        br.font.color.rgb = WHITE
        
        tb = s7.shapes.add_textbox(Inches(7.25), Inches(y), Inches(2.25), Inches(0.42))
        p = tb.text_frame.paragraphs[0]
        r1 = p.add_run()
        r1.text = f"{tag} "
        r1.font.bold = True
        r1.font.color.rgb = RED
        r1.font.size = Pt(11)
        r2 = p.add_run()
        r2.text = desc
        r2.font.color.rgb = BLACK
        r2.font.size = Pt(10)
        
    s7.notes_slide.notes_text_frame.text = (
        "Here is our user interface design. The dashboard features 6 multi-page analytical modules. Operators start at UI-01 "
        "to inspect macro KPIs and trend lines, drill down into UI-02 Pareto breakdowns, and isolate chronic outliers in UI-03."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 8: Key findings
    # ══════════════════════════════════════════════════════════════════════════
    s8 = prs.slides.add_slide(blank_layout)
    add_content_title(s8, "Key findings")
    add_footer(s8, 8)
    
    findings = [
        (0.5, "2.4x", "Weekend Failure Surge", "Evidence: Temporal heatmap reveals Saturday 18-22h failure rate hits 47.8% vs 20.0% off-peak baseline."),
        (3.55, "2.20x", "Off-Us Gateway Drop", "Evidence: Interbank interchange failure rate is 31.6% vs 14.4% domestic On-Us traffic (Z-test p < 0.001)."),
        (6.6, "100%", "Latency Timeout Cliff", "Evidence: Every transaction exceeding 3,500 ms round-trip response latency terminates in switch failure.")
    ]
    for x, num, meaning, ev in findings:
        card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.3), Inches(2.85), Inches(3.6))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BG
        
        tb_num = s8.shapes.add_textbox(Inches(x + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        p = tb_num.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = num
        r.font.name = "Cambria"
        r.font.size = Pt(38)
        r.font.bold = True
        r.font.color.rgb = RED
        
        tb_mean = s8.shapes.add_textbox(Inches(x + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        p = tb_mean.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = meaning
        r.font.name = "Calibri"
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = BLACK
        
        tb_ev = s8.shapes.add_textbox(Inches(x + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        p = tb_ev.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ev
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = GREY
        
    s8.notes_slide.notes_text_frame.text = (
        "Our analytical models uncovered three vital business findings: First, failure rates surge by 2.4x during weekend peak hours. "
        "Second, Off-Us interbank switches suffer 2.2x higher decline rates. Third, 3.5 seconds represents an absolute latency cliff."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 9: Testing and validation
    # ══════════════════════════════════════════════════════════════════════════
    s9 = prs.slides.add_slide(blank_layout)
    add_content_title(s9, "Testing and validation")
    add_footer(s9, 9)
    
    tb_score = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    p = tb_score.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "11 / 11"
    r.font.name = "Cambria"
    r.font.size = Pt(44)
    r.font.bold = True
    r.font.color.rgb = RED
    
    tb_sub = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    p = tb_sub.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "100% test cases passed\npytest: 11 / 11 · manual: 16 / 16"
    r.font.name = "Calibri"
    r.font.size = Pt(12)
    r.font.color.rgb = GREY
    
    tbl_shape = s9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.5))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(1.1)
    tbl.columns[1].width = Inches(3.9)
    tbl.columns[2].width = Inches(1.1)
    
    headers = ["Test", "Checks", "Result"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = WHITE
        
    test_rows = [
        ("TC-T-02", "Row count validation (txns >= 50k, terminals >= 200)", "Pass (0.3s)"),
        ("TC-T-04", "Referential integrity (0 orphan foreign key rows)", "Pass (0.1s)"),
        ("TC-T-07", "Anomaly logic: flag terminals > 5.0% failure SLA", "Pass (0.1s)")
    ]
    for r_idx, r_data in enumerate(test_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = tbl.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx != 1 else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.name = "Calibri"
            r.font.size = Pt(10.5)
            r.font.color.rgb = BLACK
            
    s9.notes_slide.notes_text_frame.text = (
        "Testing is fully automated with pytest. All 11 unit tests passed with 100% success rate in 0.22 seconds, covering database "
        "connectors, foreign key integrity, mathematical formulas, and interface components."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 10: Recommendation, limitations and next steps
    # ══════════════════════════════════════════════════════════════════════════
    s10 = prs.slides.add_slide(blank_layout)
    add_content_title(s10, "Recommendation, limitations and next steps")
    add_footer(s10, 10)
    
    add_badge_card(
        s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendations",
        "1. Prioritize preventive hardware overhaul on chronic terminals ATM_015, 042, 087, 123, 156.\n"
        "2. Dynamically adjust switch connection timeouts to 3,500 ms during weekend congestion windows.\n"
        "3. Enforce SLA penalty clauses on hardware vendors with failure rates > 12/1,000 txns."
    )
    add_badge_card(
        s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations",
        "1. Utilizes synthetic transaction logs; actual core switch ISO 8583 wire protocol integration was excluded.\n"
        "2. Analysis window constrained to 90 days of synthetic events.\n"
        "3. Lacks automated bi-directional API integration with enterprise ticketing systems."
    )
    add_badge_card(
        s10, 6.65, 1.25, 2.85, 3.7, "N", "Next Steps",
        "1. Implement streaming Kafka pipelines for real-time telemetry ingestion.\n"
        "2. Train machine learning classification models (XGBoost) for predictive hardware fault forecasting.\n"
        "3. Integrate automated ServiceNow dispatch webhook alerts."
    )
    s10.notes_slide.notes_text_frame.text = (
        "Based on empirical findings, we recommend targeted component servicing on chronic nodes and weekend switch load balancing. "
        "While synthetic data permitted safe simulation, next enterprise phases will introduce real-time Kafka streaming."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # SLIDE 11: DARK (Live Demo and Q&A)
    # ══════════════════════════════════════════════════════════════════════════
    s11 = prs.slides.add_slide(blank_layout)
    bg11 = s11.background
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = BLACK
    
    tb_q = s11.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(1.2))
    p = tb_q.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Live Demo & Q&A"
    r.font.name = "Cambria"
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = WHITE
    
    tb_links = s11.shapes.add_textbox(Inches(0.5), Inches(2.9), Inches(9.0), Inches(1.5))
    tf_l = tb_links.text_frame
    
    p1 = tf_l.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "App URL: [To confirm]\n"
    r1.font.name = "Calibri"
    r1.font.size = Pt(16)
    r1.font.color.rgb = PINK
    
    p2 = tf_l.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = "Code: [To confirm]\n\n"
    r2.font.name = "Calibri"
    r2.font.size = Pt(16)
    r2.font.color.rgb = PINK
    
    p3 = tf_l.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    r3 = p3.add_run()
    r3.text = "Thank You!"
    r3.font.name = "Calibri"
    r3.font.size = Pt(20)
    r3.font.bold = True
    r3.font.color.rgb = WHITE
    
    s11.notes_slide.notes_text_frame.text = (
        "Thank you esteemed mentor Mr. Dinesh Sereno and panel members. I will now switch to the live Streamlit dashboard "
        "to demonstrate interactive filtering and welcome your questions."
    )

    prs.save(output_path)
    print(f"Step 18 Presentation saved at: {output_path}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    create_presentation("docs/p096_Presentation.pptx")
