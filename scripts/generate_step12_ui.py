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

def build_step12_doc(output_path):
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
    r_sub = p_sub.add_run("Step 12 — UI Framework, Wireframes & Tailwind Design  v1.0\n")
    r_sub.font.size = Pt(13)
    r_sub.font.name = "Calibri"
    r_sub.font.color.rgb = RGBColor(74, 85, 104)
    
    meta = doc.add_table(rows=5, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    mdata = [
        ("Project & Code", "p096 | ATM Transaction Failure Analysis"),
        ("Student / Candidate", "Lavi Tewari (Roll: 2500320140085), MCA 2nd Yr DA, ABES"),
        ("Project Mentor", "Mr. Dinesh Sereno"),
        ("UI Architecture", "Streamlit Multi-Page + Tailwind CSS (via st.components.v1.html) + Plotly"),
        ("Theme Specification", "Primary: #2B6CB0 (Navy), Alert: #E53E3E (Red), Surface: #F7FAFC")
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
    
    # Section 1: Screen Inventory
    h1 = doc.add_heading("1. Screen Inventory & Architecture", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    scr_table = doc.add_table(rows=9, cols=6)
    scr_headers = ["UI ID", "Screen Name", "File Path", "Primary Components", "Active Filters", "Linked US"]
    for i, h in enumerate(scr_headers):
        scr_table.rows[0].cells[i].paragraphs[0].text = h
    scr_data = [
        ("UI-01", "Fleet Overview", "pages/01_Fleet_Overview.py", "4 Tailwind KPI cards, Plotly hourly failure line curve, summary table", "Date, Bank, Zone", "US-04, US-06"),
        ("UI-02", "Failure Category", "pages/02_Failure_Category.py", "Pareto 80/20 bar+line chart, transaction type grouped bars, donut split", "Date, Bank, Cat", "US-07, US-09"),
        ("UI-03", "Terminal Health", "pages/03_Terminal_Health.py", "Chronic outlier table (>5% SLA) with badges, MTBF column, vendor H-bar", "Date, Bank, Zone", "US-10, US-11, US-12"),
        ("UI-04", "Temporal Patterns", "pages/04_Temporal_Patterns.py", "24-hour x 7-day density heatmap, peak hour insight callout", "Date, Bank, Cat", "US-08"),
        ("UI-05", "Network & Latency", "pages/05_Network_Latency.py", "Response latency scatter plot, Pearson r callout, On-Us/Off-Us donut + Z-test", "Date, Bank, Card", "US-13, US-14"),
        ("UI-06", "Financial Impact", "pages/06_Financial_Impact.py", "INR at Risk KPI cards, failure category treemap, transaction value breakdown", "Date, Bank, Type", "US-06"),
        ("UI-07", "Data Quality", "app.py (Sidebar)", "Data quality audit box, missing count, duplicate counter, outlier badge", "None (Global)", "US-03"),
        ("UI-08", "Data Export", "app.py (Bottom Tab)", "st.download_button, CSV stream generator, row count badge", "All Active Filters", "US-15")
    ]
    for r_idx, r_data in enumerate(scr_data, start=1):
        for c_idx, text in enumerate(r_data):
            scr_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(scr_table, [0.6, 1.2, 1.7, 1.5, 0.9, 0.6])
    doc.add_paragraph()
    
    # Section 2: Wireframes
    h1 = doc.add_heading("2. ASCII Screen Wireframes", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    wireframe_text = """
UI-01: FLEET OVERVIEW
+---------------------------------------------------------------------------------------+
|  ATM p096: ATM Transaction Failure Analysis      [Overview] [Taxonomy] [Terminals]... |
+-----------------+---------------------------------------------------------------------+
| SIDEBAR FILTERS |  TAILWIND KPI ROW                                                   |
| [Date: Jan-Mar] |  +------------+  +------------+  +------------+  +----------------+ |
| [Bank: All    ] |  | Total Txns |  | Approved   |  | Failed     |  | Disrupted INR  | |
| [Zone: Metro  ] |  |   50,003   |  | 39,347(78%)|  | 10,656(21%)|  | ₹ 4.28 Crore   | |
| [Tx Type: All ] |  +------------+  +------------+  +------------+  +----------------+ |
|                 |                                                                     |
| [Quality Audit] |  HOURLY FAILURE RATE TREND (Plotly Interactive Line Chart)          |
| Missing: 1,000  |  +----------------------------------------------------------------+ |
| Duplicates: 3   |  | Y: Failure Rate %  (Spikes at 20:00 Saturday & 11:00 Sunday)   | |
|                 |  | X: Timestamp (Hourly Window)                                   | |
| [Reset Filters] |  +----------------------------------------------------------------+ |
|                 |  INSIGHT BOX: Peak fleet failure is concentrated in Saturday evening|
+-----------------+---------------------------------------------------------------------+

UI-03: TERMINAL HEALTH & OUTLIER RANKING
+---------------------------------------------------------------------------------------+
|  ATM p096: ATM Transaction Failure Analysis  >  Terminal Health Diagnostics           |
+-----------------+---------------------------------------------------------------------+
| SIDEBAR FILTERS |  CHRONIC PROBLEM TERMINALS (>5.0% SLA Threshold)                    |
| [Bank: All    ] |  +-------------+----------+-----------+------------+-------+------+ |
| [Zone: All    ] |  | Terminal ID | Location | City      | Fail Rate  | MTBF  | Stat | |
| [Vendor: NCR  ] |  +-------------+----------+-----------+------------+-------+------+ |
|                 |  | ATM_015     | Andheri  | Mumbai    | [!] 50.6%  | 4.2 h | CRIT | |
|                 |  | ATM_042     | CP       | Delhi     | [!] 44.7%  | 4.8 h | CRIT | |
|                 |  | ATM_123     | Koramang | Bengaluru | [!] 49.5%  | 4.1 h | CRIT | |
|                 |  +-------------+----------+-----------+------------+-------+------+ |
|                 |                                                                     |
|                 |  HARDWARE VENDOR BENCHMARK (Normalised Failures per 1,000 Txns)     |
|                 |  NCR      : [==========================] 14.7%                      |
|                 |  Diebold  : [====================] 11.2%                            |
|                 |  Hitachi  : [==================] 10.8%                              |
+-----------------+---------------------------------------------------------------------+
"""
    p_wf = doc.add_paragraph()
    r_wf = p_wf.add_run(wireframe_text)
    r_wf.font.name = "Courier New"
    r_wf.font.size = Pt(8)
    r_wf.font.color.rgb = RGBColor(45, 55, 72)
    doc.add_paragraph()
    
    # Section 3: Global Layout
    h1 = doc.add_heading("3. Global Layout & Theme Specifications", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    doc.add_paragraph("• Palette: Primary Navy (#2B6CB0), Danger Alert Red (#E53E3E), Success Forest Green (#38A169), Neutral Canvas (#F7FAFC).")
    doc.add_paragraph("• Typography: Google Fonts Inter (sans-serif) loaded via Tailwind CDN; clean fallback to Calibri.")
    doc.add_paragraph("• State Management: st.session_state maintains active filter criteria across tab navigations without data reload.")
    doc.add_paragraph("• Responsive Grid: Wide-screen default (layout='wide') with CSS flex-wrap cards supporting tablets and laptops.")
    doc.add_paragraph()
    
    # Section 4: Reusable Component Catalog
    h1 = doc.add_heading("4. Reusable Component Catalog", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    comp_table = doc.add_table(rows=6, cols=4)
    comp_headers = ["Component Name", "Tailwind Classes", "Streamlit Rendering Method", "Usage Scope"]
    for i, h in enumerate(comp_headers):
        comp_table.rows[0].cells[i].paragraphs[0].text = h
    comp_data = [
        ("Tailwind Top Header", "bg-blue-900 text-white rounded-xl shadow-lg px-6 py-3", "components.html(html, height=100)", "Global (All Pages)"),
        ("Executive KPI Card", "bg-white rounded-xl shadow-md p-4 border-l-4 border-blue-600", "components.html(html, height=125)", "UI-01, UI-06"),
        ("Diagnostic Insight Box", "bg-blue-50 border-l-4 border-blue-500 rounded-r-xl px-4 py-3", "components.html(html, height=65)", "UI-01, UI-02, UI-04, UI-05"),
        ("Chronic Outlier Badge", "bg-red-100 text-red-800 rounded-full px-3 py-1 font-bold", "st.markdown(unsafe_allow_html=True)", "UI-03"),
        ("Data Quality Warning", "bg-yellow-100 border-l-4 border-yellow-500 text-yellow-800", "st.sidebar.markdown(unsafe_allow_html=True)", "UI-07 (Sidebar)")
    ]
    for r_idx, r_data in enumerate(comp_data, start=1):
        for c_idx, text in enumerate(r_data):
            comp_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(comp_table, [1.5, 2.0, 1.8, 1.2])
    doc.add_paragraph()
    
    # Section 5: Best Chart Per OQ
    h1 = doc.add_heading("5. Optimal Visual Chart Matrix per Operational Question", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    
    chart_table = doc.add_table(rows=11, cols=5)
    chart_headers = ["OQ ID", "Best Chart Type", "Plotly Function", "Analytical Justification", "Visual Dimensions"]
    for i, h in enumerate(chart_headers):
        chart_table.rows[0].cells[i].paragraphs[0].text = h
    chart_data = [
        ("OQ-01", "Time Series Line", "px.line()", "Reveals intraday surges and macro temporal direction", "X: Timestamp, Y: Failure Rate %"),
        ("OQ-02", "Pareto Bar + Line", "go.Bar() + go.Scatter()", "Applies 80/20 rule to isolate vital few root causes", "Bars: Counts, Line: Cumulative %"),
        ("OQ-03", "Ranked Horizontal Bar", "px.bar(orientation='h')", "Sorts outlier terminals exceeding 5% SLA threshold", "Y: Terminal ID, X: Failure Rate %"),
        ("OQ-04", "Temporal 2D Heatmap", "px.density_heatmap()", "Visualizes 24x7 congestion matrix without clutter", "X: Hour, Y: Day of Week, Z: Count"),
        ("OQ-05", "Grouped Categorical Bar", "px.bar(barmode='group')", "Direct comparison across Cash vs PIN vs Inquiry", "X: Tx Type, Y: Fail Rate, Color: Status"),
        ("OQ-06", "Continuous Scatter", "px.scatter()", "Shows latency cluster threshold driving hard timeouts", "X: Response Time (ms), Y: Outcome"),
        ("OQ-07", "Donut Proportion Chart", "px.pie(hole=0.4)", "Part-to-whole interbank gateway share", "Values: Count, Names: Card Network"),
        ("OQ-08", "Hierarchical Treemap", "px.treemap()", "Area corresponds directly to Rupee financial exposure", "Hierarchy: Category > Type, Area: Amount"),
        ("OQ-09", "Normalized Bar Chart", "px.bar()", "Controls for fleet size differences across vendors", "X: Vendor, Y: Failures / 1,000 Txns"),
        ("OQ-10", "Ascending Bar Chart", "px.bar()", "Isolates terminals needing immediate preventative PM", "X: Terminal ID, Y: MTBF Running Hours")
    ]
    for r_idx, r_data in enumerate(chart_data, start=1):
        for c_idx, text in enumerate(r_data):
            chart_table.rows[r_idx].cells[c_idx].paragraphs[0].text = text
    style_table(chart_table, [0.7, 1.4, 1.4, 1.8, 1.2])
    doc.add_paragraph()
    
    # Section 6: RTM Traceability
    h1 = doc.add_heading("6. Traceability Mapping: UI Screens to Business Requirements", level=1)
    h1.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    add_callout(doc,
        "UI TRACEABILITY SIGN-OFF: Screens UI-01 through UI-08 fully satisfy every requirement in BR-01 through BR-12. "
        "All 10 Operational Questions have verified visualization mappings. Tailwind components are production-ready in src/ui.py.",
        title="UI SPECIFICATION APPROVED")
        
    doc.save(output_path)
    print(f"Step 12 DOCX created at: {output_path}")

if __name__ == "__main__":
    os.makedirs("docs", exist_ok=True)
    build_step12_doc("docs/12_UI_Framework.docx")
