"""
Check what sheets exist in KB25019 Excel file
"""

import openpyxl


def check_kb25019_sheets():
    """Check what sheets are in the KB25019 Excel file."""
    
    excel_file = "../CT&INV&PL KB25019 DAP(1)(1).xlsx"
    
    try:
        workbook = openpyxl.load_workbook(excel_file, data_only=False)
        
        print(f"📋 Sheets in KB25019:")
        for i, sheet_name in enumerate(workbook.sheetnames, 1):
            print(f"   {i}. {sheet_name}")
        
        workbook.close()
        
    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    check_kb25019_sheets()
