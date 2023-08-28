import sys, os, subprocess
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def is_empty_cell(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            xml_content = run._r.xml
            
            if xml_content.strip(): # Checks for any content - fix for image returning empty
                return False  
    return True

def comment_formatting(paragraph, finding_text):
    paragraph.add_run(finding_text).bold = True
    paragraph.runs[-1].font.color.rgb = RGBColor(0xFF, 0, 0)

def mark_empty_cells(cell):
    paragraph = cell.paragraphs[0]

    if is_empty_cell(cell):
        finding_text = "EMPTY"
        comment_formatting(paragraph, finding_text)

def check_and_mark_cdash_cells(cell):
    cdash_text = cell.text.strip()
    paragraph = cell.paragraphs[0]
    
    if not check_for_existing_findings(cell.paragraphs):
        if "-" in cdash_text and (" -" in cdash_text or "- " in cdash_text):
            finding_text = "\nREDUNDANT SPACES"
            comment_formatting(paragraph, finding_text)
        
        if "—" in cdash_text:
            finding_text = "\nDASH INSTEAD OF HYPHEN"
            comment_formatting(paragraph, finding_text)
        
        if "-" not in cdash_text and "N/A" not in cdash_text and "—" not in cdash_text:
            finding_text = "\nMISSING HYPHEN"
            comment_formatting(paragraph, finding_text)
        
        return False

def find_cdash_column(table):
    for col_idx, cell in enumerate(table.rows[0].cells):
        if "CDASH" in cell.text.upper():
            return col_idx
    return None

def check_and_mark_alignment_issue(cell, last_column_cell):
    alignment_missing = False
    
    for paragraph in cell.paragraphs:
        alignment = paragraph.alignment
        paragraph = last_column_cell.paragraphs[0]

        if alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            if "center aligned" not in last_column_cell.text and not check_for_existing_findings(last_column_cell.paragraphs):
                finding_text = "\nMISSING ALIGNMENT/FORMATTING COMMENT"
                comment_formatting(paragraph, finding_text)
                alignment_missing = True
                break
    
    return alignment_missing

def check_for_existing_findings(paragraphs):
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if run.bold and run.font.color.rgb == RGBColor(0xFF, 0, 0):
                text = run.text.strip()
                if text.isupper():
                    run.font.underline = True
                    return True
    return False

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 word_table_scanner.py <box_file_link_or_local_path>") # link args not accepted for now
    else:
        input_arg = sys.argv[1]
        
        if os.path.exists(input_arg):
            file_path = input_arg
        else:
            file_path = download_box_file(input_arg) # currently not functional 
        
        if file_path:
            document = Document(file_path)
                
            modified = False
            
            for table in document.tables:
                cdash_col_idx = find_cdash_column(table)
                for row_idx, row in enumerate(table.rows):

                    alignment_issue_cell = row.cells[-1]  
                
                    for col_idx, cell in enumerate(row.cells[:-1], start=1):  # Exclude the last column
                        if check_for_existing_findings(cell.paragraphs):
                            modified = True

                        alignment_issue = check_and_mark_alignment_issue(cell, alignment_issue_cell)

                        if alignment_issue:
                            modified = True
                            break 
                
                if cdash_col_idx is not None:
                    for row_idx, row in enumerate(table.rows[1:], start=1):
                        cdash_cell = row.cells[cdash_col_idx]
                        
                        if check_and_mark_cdash_cells(cdash_cell):
                            modified = True
                            break  
                
                for row_idx, row in enumerate(table.rows):
                    for col_idx, cell in enumerate(row.cells):

                        if is_empty_cell(cell):
                            mark_empty_cells_with(cell)
                            modified = True
            
            if modified:
                document.save(file_path)  
                print("\033[91mFindings recorded in the file.\033[0m")

                try:
                    subprocess.run(["open", file_path], check=True)
                except subprocess.CalledProcessError:
                    print("Failed to open the modified file.")
            else:
                print("\033[92mNo findings found!\033[0m")
