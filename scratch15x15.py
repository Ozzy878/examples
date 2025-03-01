"""scratch15x15.py
Expand an existing 10x10 multiplication table and resize columns"""
import win32com.client as win32
from pathlib import Path

output_file_name = "Scratch_15x15"
output_folder = "S://GitHub/examples/output_files"
output_file = Path('{}/{}.xlsx'.format(output_folder, output_file_name))

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

wb = excel.Workbooks.Add()
ws = wb.Worksheets('Sheet1')

# Initialize the first row and column
ws.Range("B2:P2").Value = [i for i in range(1, 16)]
ws.Range("B2:B16").Value = list(zip([i for i in range(1, 16)]))

# Populate the table
ws.Range("C3").Formula = "=$B3*C$2"
ws.Range("C3:C3").Select()
excel.Selection.AutoFill(ws.Range("C3:P3"), win32.constants.xlFillDefault)
ws.Range("C3:P3").Select()
excel.Selection.AutoFill(ws.Range("C3:P16"), win32.constants.xlFillDefault)

# Save the spreadsheet
wb.SaveAs(str(output_file), FileFormat=51, ConflictResolution=2)
excel.Application.Quit()
