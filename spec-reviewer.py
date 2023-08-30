import sys, os, subprocess
import argparse
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def parse_arguments():
    parser = argparse.ArgumentParser(description="Spec Scanner with Optional Checks")
    parser.add_argument("file_path", help="Path to the Word document file")
    parser.add_argument("--ignore-line-breaks", action="store_true", help="Ignore line break check")
    parser.add_argument("--ignore-formatting", action="store_true", help="Ignore formatting check")
    return parser.parse_args()

def is_empty_cell(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            xml_content = run._r.xml
            if xml_content.strip(): 
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
        return True
    return False

def check_and_mark_cdash_cells(cell):
    cdash_text = cell.text.strip()
    paragraph = cell.paragraphs[0]
    if not check_for_existing_findings(cell.paragraphs):
        if "-" in cdash_text and (" -" in cdash_text or "- " in cdash_text):
            finding_text = "\nREDUNDANT SPACES"
            comment_formatting(paragraph, finding_text)
            return True
        
        if "—" in cdash_text:
            finding_text = "\nDASH INSTEAD OF HYPHEN"
            comment_formatting(paragraph, finding_text)
            return True
        
        if "-" not in cdash_text and "N/A" not in cdash_text:
            finding_text = "\nMISSING HYPHEN"
            comment_formatting(paragraph, finding_text)
            return True
        
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

        if alignment == WD_PARAGRAPH_ALIGNMENT.CENTER and not ignore_formatting:
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

def check_line_breaks(cell):
    if not ignore_line_breaks and '\n' in cell.text and not check_for_existing_findings(cell.paragraphs):
        paragraph = cell.paragraphs[-1]
        finding_text = "\nCHECK LINE BREAKS"
        comment_formatting(paragraph, finding_text)
        return True
    return False
  
if __name__ == "__main__":
    args = parse_arguments()
    file_path = args.file_path
    ignore_line_breaks = args.ignore_line_breaks
    ignore_formatting = args.ignore_formatting

    if len(sys.argv) < 2:
        print("Usage: python3 word_table_scanner.py <local_file_path> optional-ignore-flag")
    else:
        input_arg = sys.argv[1]
        
        if os.path.exists(input_arg):
            file_path = input_arg
        
        if file_path:
            document = Document(file_path)
                
            modified = False
            
            for table in document.tables:
                cdash_col_idx = find_cdash_column(table)
                for row_idx, row in enumerate(table.rows[1:], start=1):
                    alignment_issue_cell = row.cells[-1]
                    
                    if cdash_col_idx is not None:
                        cdash_cell = row.cells[cdash_col_idx]
                        
                        if check_and_mark_cdash_cells(cdash_cell):
                            modified = True
                            break

                    for col_idx, cell in enumerate(row.cells[:-1], start=1):  # Exclude the last column
                        if check_for_existing_findings(cell.paragraphs):
                                modified = True

                        if check_line_breaks(cell):
                            modified = True

                        alignment_issue = check_and_mark_alignment_issue(cell, alignment_issue_cell)

                        if alignment_issue:
                            modified = True
                            break   
                
                for row_idx, row in enumerate(table.rows):
                    for col_idx, cell in enumerate(row.cells):
                        if is_empty_cell(cell):
                            mark_empty_cells(cell)
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
