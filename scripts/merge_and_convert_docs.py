import os
import win32com.client
from pathlib import Path

def merge_docx_and_export_pdf():
    base_dir = Path(__file__).resolve().parent.parent
    docs_dir = base_dir / "docs"
    final_dir = base_dir / "final_submission"
    final_dir.mkdir(parents=True, exist_ok=True)
    
    # List of docx files to merge in logical sequence
    doc_files = [
        docs_dir / "p096_Project_Document.docx",
        docs_dir / "01_Project_Understanding.docx",
        docs_dir / "02_HMW_Statements.docx",
        docs_dir / "03_Needs_Analysis.docx",
        docs_dir / "05_BRD.docx",
        docs_dir / "07_TRD.docx",
        docs_dir / "09_ER_Diagram.docx",
        docs_dir / "11_User_Stories.docx",
        docs_dir / "12_UI_Framework.docx",
        docs_dir / "13_Test_Cases.docx",
        docs_dir / "14_BUILD_PLAN.docx",
    ]
    
    print("Files to merge:")
    for f in doc_files:
        print(" -", f.name, "exists:", f.exists())
        
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    
    try:
        # Open the master document as the base
        master_path = str(doc_files[0])
        doc = word.Documents.Open(master_path)
        
        # Word constants
        wdStory = 6
        wdSectionBreakNextPage = 2
        wdPageBreak = 7
        
        # Move to end of master document
        for f in doc_files[1:]:
            if not f.exists():
                print(f"Skipping {f.name} (not found)")
                continue
                
            selection = word.Selection
            selection.EndKey(Unit=wdStory)
            selection.InsertBreak(Type=wdPageBreak)
            
            # Insert file content
            print(f"Appending {f.name}...")
            selection.InsertFile(FileName=str(f))
            
        merged_docx = final_dir / "project_document.docx"
        merged_pdf = final_dir / "project_document.pdf"
        
        print(f"Saving merged document to {merged_docx}...")
        doc.SaveAs2(str(merged_docx))
        
        print(f"Exporting to PDF: {merged_pdf}...")
        wdFormatPDF = 17
        doc.SaveAs2(str(merged_pdf), FileFormat=wdFormatPDF)
        
        doc.Close(False)
        print("Successfully merged and exported project_document.pdf!")
        
    except Exception as e:
        print(f"Error during merge: {e}")
        raise
    finally:
        word.Quit()

if __name__ == "__main__":
    merge_docx_and_export_pdf()
