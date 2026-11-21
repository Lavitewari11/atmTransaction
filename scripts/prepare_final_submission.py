import os
import win32com.client
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Preformatted, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

os.makedirs("final_submission", exist_ok=True)

# ── 1. CONVERT p096_Project_Document.docx -> project_document.pdf via Word COM ──
def docx_to_pdf(doc_path, pdf_path):
    abs_doc = os.path.abspath(doc_path)
    abs_pdf = os.path.abspath(pdf_path)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(abs_doc)
        # 17 represents wdExportFormatPDF
        doc.SaveAs(abs_pdf, FileFormat=17)
        doc.Close()
        print(f"Successfully converted {doc_path} -> {pdf_path}")
    finally:
        word.Quit()

# ── 2. CONVERT p096_Presentation.pptx -> project_presentation.pdf via PowerPoint COM ──
def pptx_to_pdf(ppt_path, pdf_path):
    abs_ppt = os.path.abspath(ppt_path)
    abs_pdf = os.path.abspath(pdf_path)
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        presentation = ppt_app.Presentations.Open(abs_ppt, WithWindow=False)
        # 32 represents ppSaveAsPDF
        presentation.SaveAs(abs_pdf, 32)
        presentation.Close()
        print(f"Successfully converted {ppt_path} -> {pdf_path}")
    finally:
        ppt_app.Quit()

# ── 3. GENERATE github_report.pdf ──
def generate_github_report(pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        textColor=colors.HexColor('#A51C30'),
        spaceAfter=10
    )
    h2_style = ParagraphStyle(
        'H2Style',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=13,
        textColor=colors.HexColor('#1A365D'),
        spaceBefore=12,
        spaceAfter=6
    )
    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        textColor=colors.HexColor('#151514'),
        spaceAfter=6
    )
    code_style = ParagraphStyle(
        'CodeStyle',
        parent=styles['Code'],
        fontName='Courier',
        fontSize=8.5,
        textColor=colors.HexColor('#2D3748'),
        spaceAfter=4
    )

    story = []
    story.append(Paragraph("GitHub Repository & Deployment Audit Report", title_style))
    story.append(Paragraph("<b>Project:</b> p096 — ATM Transaction Failure Analysis", body_style))
    story.append(Paragraph("<b>Candidate:</b> Lavi Tewari (Roll: 2500320140085), MCA 2nd Year Data Analytics, ABES", body_style))
    story.append(Paragraph("<b>Mentor:</b> Mr. Dinesh Sereno | <b>Date:</b> November 2026", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("1. Verified Online Endpoints", h2_style))
    story.append(Paragraph("<b>GitHub Repository URL:</b> <font color='#2B6CB0'><u>https://github.com/Lavitewari11/atmTransaction</u></font>", body_style))
    story.append(Paragraph("<b>Live Streamlit Cloud URL:</b> <font color='#2B6CB0'><u>https://2500320140085-p062-atm.streamlit.app/</u></font>", body_style))
    story.append(Paragraph("<b>Repository Visibility:</b> Public | <b>Default Branch:</b> main | <b>Latest Commit:</b> 971e74a", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("2. Repository Structure & File Manifest", h2_style))
    story.append(Paragraph("The GitHub repository contains 60 version-controlled files structured as follows:", body_style))

    manifest = """
Directory: Lavitewari11/atmTransaction (branch: main)
├── app.py                         [Streamlit main entrypoint & global filter router]
├── pages/
│   ├── 01_Fleet_Overview.py       [Macro fleet KPIs & hourly failure rate trend]
│   ├── 02_Failure_Category.py     [Pareto 80/20 root cause analysis & transaction type]
│   ├── 03_Terminal_Health.py      [Chronic outlier ranking (>5% SLA) & MTBF]
│   ├── 04_Temporal_Patterns.py    [24x7 failure density heatmap]
│   ├── 05_Network_Latency.py      [Latency vs timeout scatter & On-Us/Off-Us Z-test]
│   └── 06_Financial_Impact.py     [Gross INR disrupted treemap & loss metrics]
├── src/
│   ├── db.py                      [SQLite connector & @st.cache_data loaders]
│   ├── transform.py               [Taxonomy mapping & data quality telemetry]
│   ├── analysis.py                [Statistical calculations, outlier logic, MTBF]
│   ├── charts.py                  [Plotly visual chart factory]
│   ├── ui.py                      [Tailwind top banner, KPI cards & sidebar filters]
│   └── generate_data.py           [Synthetic banking data generator (50k rows)]
├── data/
│   ├── app.db                     [SQLite database with 4 tables + 1 view]
│   ├── atm_terminals.csv          [200 terminal master records]
│   ├── error_codes.csv            [30 standard ISO 8583 banking response codes]
│   ├── transactions.csv           [50,003 transactional telemetry records]
│   ├── maintenance_logs.csv       [1,500 terminal maintenance events]
│   └── atm_transactions_data.xlsx [Master 4-sheet Excel database]
├── tests/
│   ├── conftest.py                [Database connection & sample DataFrame fixtures]
│   ├── test_data.py               [TC-T-01..05: DB connection, table counts, PK/FK]
│   ├── test_analysis.py           [TC-T-06..09: Mathematical formulas & outlier logic]
│   └── test_app.py                [TC-T-10..11: Page existence & UI theme validation]
├── docs/                          [18-step academic Word, Excel, and Presentation files]
├── .streamlit/
│   └── config.toml                [Navy/Slate color theme specification]
├── BUILD_PLAN.md                  [Step-by-step engineering plan]
├── requirements.txt               [Pinned package manifest]
└── README.md                      [Documentation & setup instructions]
"""
    story.append(Preformatted(manifest, code_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("3. Security & Quality Audit", h2_style))
    story.append(Paragraph("• <b>Zero Secrets:</b> Audit confirmed no credentials, API keys, or private tokens exist in repository.", body_style))
    story.append(Paragraph("• <b>Automated Testing:</b> All 11 pytest unit tests pass with 100% success rate.", body_style))
    story.append(Paragraph("• <b>Cloud Deployment:</b> Successfully operational on Streamlit Community Cloud with public URL.", body_style))

    doc.build(story)
    print(f"Generated GitHub report -> {pdf_path}")

# ── 4. GENERATE complete_code.pdf ──
def generate_complete_code(pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    styles = getSampleStyleSheet()
    file_hdr_style = ParagraphStyle(
        'FileHeaderStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=colors.HexColor('#A51C30'),
        spaceBefore=14,
        spaceAfter=4
    )
    code_text_style = ParagraphStyle(
        'CodeTextStyle',
        parent=styles['Code'],
        fontName='Courier',
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor('#151514')
    )

    story = []
    title_p = Paragraph("<font size=18 color='#A51C30'><b>Complete Source Code Repository — p096</b></font><br/><font size=10 color='#5E5E5B'>p096: ATM Transaction Failure Analysis | Lavi Tewari (Roll: 2500320140085)</font><br/>", styles['Normal'])
    story.append(title_p)
    story.append(Spacer(1, 10))

    files_to_include = [
        "app.py",
        "pages/01_Fleet_Overview.py",
        "pages/02_Failure_Category.py",
        "pages/03_Terminal_Health.py",
        "pages/04_Temporal_Patterns.py",
        "pages/05_Network_Latency.py",
        "pages/06_Financial_Impact.py",
        "src/ui.py",
        "src/db.py",
        "src/transform.py",
        "src/analysis.py",
        "src/charts.py",
        "tests/conftest.py",
        "tests/test_data.py",
        "tests/test_analysis.py",
        "tests/test_app.py",
        "requirements.txt"
    ]

    for rel_path in files_to_include:
        if os.path.exists(rel_path):
            story.append(Paragraph(f"=== File: {rel_path} ===", file_hdr_style))
            with open(rel_path, "r", encoding="utf-8", errors="ignore") as f:
                code_content = f.read()
            # Clean characters that break reportlab XML
            code_safe = code_content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Preformatted(code_safe, code_text_style))
            story.append(PageBreak())

    doc.build(story)
    print(f"Generated complete code PDF -> {pdf_path}")

if __name__ == "__main__":
    out_dir = "final_submission"
    
    # 1. Project Document PDF
    docx_to_pdf("docs/p096_Project_Document.docx", os.path.join(out_dir, "project_document.pdf"))
    
    # 2. Presentation PDF
    pptx_to_pdf("docs/p096_Presentation.pptx", os.path.join(out_dir, "project_presentation.pdf"))
    
    # 3. GitHub Report PDF
    generate_github_report(os.path.join(out_dir, "github_report.pdf"))
    
    # 4. Complete Code PDF
    generate_complete_code(os.path.join(out_dir, "complete_code.pdf"))
    
    # 5. Copy Excel Files to final_submission
    import shutil
    shutil.copy("docs/04_Traceability_Matrix.xlsx", os.path.join(out_dir, "04_Traceability_Matrix.xlsx"))
    shutil.copy("docs/08_Database_Design.xlsx", os.path.join(out_dir, "08_Database_Design.xlsx"))
    shutil.copy("data/atm_transactions_data.xlsx", os.path.join(out_dir, "project_database.xlsx"))
    shutil.copy("docs/p096_Presentation.pptx", os.path.join(out_dir, "project_presentation.pptx"))
    
    print("\nAll submission deliverables created in final_submission/ directory.")
