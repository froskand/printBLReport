import os
import sys
import win32com.client
from tkinter import messagebox
import tkinter as tk


def export_excel_to_pdf(workbook_path):
    """
    Export Excel sheets to PDF with header repetition and page numbering.
    
    Args:
        workbook_path: Full path to the Excel workbook
    
    Returns:
        bool: True if successful, False otherwise
    """
    excel = None
    wb = None
    
    try:
        # Initialize COM object for Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Run in background
        excel.DisplayAlerts = False  # Suppress Excel alerts
        
        # Open workbook
        print(f"Opening workbook: {workbook_path}")
        wb = excel.Workbooks.Open(os.path.abspath(workbook_path))

        
        # Get document information from Introduction sheet
        intro_sheet = wb.Worksheets("Introduction")
        doc_id = str(intro_sheet.Range("C5").Value)
        doc_rev = str(intro_sheet.Range("C6").Value)
        
        print(f"Document ID: {doc_id}")
        print(f"Document Revision: {doc_rev}")
        
        # Build PDF filename
        workbook_dir = os.path.dirname(workbook_path)
        pdf_filename = os.path.join(
            workbook_dir, 
            f"{doc_id} rev {doc_rev} BaselineReport.pdf"
        )
        
        # Check if file already exists
        if os.path.exists(pdf_filename):
            root = tk.Tk()
            root.withdraw()
            response = messagebox.askyesno(
                "File Exists",
                f"A file with this name already exists:\n\n"
                f"{pdf_filename}\n\n"
                f"Do you want to overwrite it?",
                icon='warning'
            )
            root.destroy()
            
            if not response:
                print("User cancelled - file already exists")
                messagebox.showinfo("Cancelled", "PDF save cancelled.")
                return False
        
        # Confirm save
        root = tk.Tk()
        root.withdraw()
        response = messagebox.askokcancel(
            "Save PDF",
            f"Do you want to save the baseline document as a PDF?\n\n"
            f"Saving as: {pdf_filename}\n"
            f"Path length: {len(pdf_filename)} characters"
        )
        root.destroy()
        
        if not response:
            print("User cancelled save operation")
            return False
        
        # Set up repeating rows and page numbering for Baseline content sheet
        print("Configuring Baseline content page setup...")
        baseline_sheet = wb.Worksheets("Baseline content")
        baseline_sheet.PageSetup.PrintTitleRows = "$2:$4"  # Repeat row 2:4

        # Set up repeating rows and page numbering for Review sheet
        print("Configuring Review page setup...")
        review_sheet = wb.Worksheets("Review")
        review_sheet.PageSetup.PrintTitleRows = "$24:$25"  # Repeat row 24:25

        
        # Select sheets to export
        sheets_to_export = [
            "Introduction", 
            "Review", 
            "Baseline content", 
            "Change tracking"
        ]
        
        print(f"Selecting sheets: {', '.join(sheets_to_export)}")
        wb.Worksheets(sheets_to_export).Select()
        
        # Export to PDF
        pdf_filename = os.path.abspath(pdf_filename)
        print(f"Exporting to PDF: {pdf_filename}")
        wb.ActiveSheet.ExportAsFixedFormat(
            Type=0,  # xlTypePDF = 0
            Filename=pdf_filename,
            Quality=0,  # xlQualityStandard = 0
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        
        # Verify file was created
        if not os.path.exists(pdf_filename):
            messagebox.showerror(
                "Save Failed",
                f"Error: PDF file was not saved successfully!\n\n"
                f"Expected location: {pdf_filename}"
            )
            print("ERROR: PDF file was not created")
            return False
        else:
            file_size = os.path.getsize(pdf_filename)
            file_size_mb = file_size / (1024 * 1024)
            
            messagebox.showinfo(
                "Success",
                f"PDF saved successfully!\n\n"
                f"{pdf_filename}\n\n"
                f"File size: {file_size_mb:.2f} MB"
            )
            print(f"SUCCESS: PDF created ({file_size_mb:.2f} MB)")
            return True
        
    except Exception as e:
        error_msg = f"Failed to save PDF: {str(e)}"
        print(f"ERROR: {error_msg}")
        messagebox.showerror("Error", error_msg)
        return False
        
    finally:
        # Clean up - close workbook and Excel
        try:
            if wb:
                intro_sheet = wb.Worksheets("Introduction")
                intro_sheet.Select()
                wb.Close(SaveChanges=False)
                print("Workbook closed")
        except:
            pass
            
        try:
            if excel:
                excel.Quit()
                print("Excel application closed")
        except:
            pass


def main():
    """
    Main function to run the script.
    """
    # Check if workbook path is provided as command line argument
    if len(sys.argv) > 1:
        workbook_path = sys.argv[1]
    else:
        # Default path - UPDATE THIS to your workbook location
        workbook_path = r"C:\path\to\your\workbook.xlsx"
        
        # Or use file dialog to select workbook
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        workbook_path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        root.destroy()
        
        if not workbook_path:
            print("No file selected. Exiting.")
            return
    
    # Check if file exists
    if not os.path.exists(workbook_path):
        print(f"ERROR: File not found: {workbook_path}")
        messagebox.showerror("File Not Found", f"Could not find file:\n{workbook_path}")
        return
    
    print(f"Starting PDF export process...")
    print(f"Workbook: {workbook_path}")
    
    # Run the export
    success = export_excel_to_pdf(workbook_path)
    
    if success:
        print("\n=== Export completed successfully ===")
    else:
        print("\n=== Export failed or was cancelled ===")


if __name__ == "__main__":
    main()