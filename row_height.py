"""Set row heights and align text within the cell"""
import win32com.client as win32
from pathlib import Path

output_file_name = "row_height_format"
output_folder = "S://GitHub/examples/output_files"
output_file = Path('{}/{}.xlsx'.format(output_folder, output_file_name))

excel = win32.gencache.EnsureDispatch('Excel.Application')

wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

ws.Range("A1:A2").Value = "1 line"
ws.Range("B1:B2").Value = "Two\nlines"
ws.Range("C1:C2").Value = "Three\nlines\nhere"
ws.Range("D1:D2").Value = "This\nis\nfour\nlines"
ws.Rows(1).RowHeight = 60
ws.Range("2:2").RowHeight = 120

ws.Rows(1).VerticalAlignment = win32.constants.xlCenter
ws.Range("2:2").VerticalAlignment = win32.constants.xlCenter

# Alternately, you can autofit all rows in the worksheet
# ws.Rows.AutoFit()

wb.SaveAs(str(output_file), FileFormat=51, ConflictResolution=2)
excel.Application.Quit()
