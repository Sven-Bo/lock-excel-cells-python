from pathlib import Path  # standard python module

import win32com.client  # pip install pywin32


# Helper Function to convert RGB to integers
def rgb_to_int(rgb):
    color_int = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return color_int


# ---------------- SETTINGS ----------------
excel_file_name = "EBIT_2022.xlsx"
worksheet_password = "pw"
lock_color = rgb_to_int((255, 255, 153))
# ------------------------------------------


# get current dir (using cwd() is needed when using notebooks)
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
excel_file_path = current_dir / excel_file_name
output_path = current_dir / f"{excel_file_path.stem}_locked{excel_file_path.suffix}"

# launch Excel & open workbook
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
wb = xl.Workbooks.Open(excel_file_path)

# Within each sheet, iterate over used range & lock cells
for sh in wb.Sheets:
    wb.Worksheets(sh.Name).Activate()
    ws = xl.ActiveSheet

    rowcount = ws.UsedRange.Rows.Count
    colcount = ws.UsedRange.Columns.Count

    ws.Unprotect(Password=worksheet_password)
    for row in range(1, rowcount + 1):
        for col in range(1, colcount + 1):
            cell = ws.Cells(row, col)

            #             if cell.Font.FontStyle == "Bold":
            #                 cell.Locked = False

            #             if cell.Value == "Input Fields":
            #                 cell.Locked = False

            #             if not cell.HasFormula:
            #                 cell.Locked = False

            if cell.Interior.color == lock_color:
                cell.Locked = False

            else:
                cell.Locked = True
    ws.Protect(Password=worksheet_password)

# Save & close workbook
wb.SaveAs(Filename=str(output_path))
wb.Close(False)

# Quit Excel instance
xl.Quit()
xl = None
