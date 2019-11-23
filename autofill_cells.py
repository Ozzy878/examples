"""Autofill cell contents"""
import win32com.client as win32
from pathlib import Path

output_file_name = "autofill_cells"
output_folder = "S://GitHub/examples/output_files"
output_file = Path('{}/{}.xlsx'.format(output_folder, output_file_name))

excel = win32.gencache.EnsureDispatch('Excel.Application')

wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

ws.Range("A1").Value = 1
ws.Range("A2").Value = 2
ws.Range("A1:A2").AutoFill(ws.Range("A1:A10"), win32.constants.xlFillDefault)

wb.SaveAs(str(output_file), FileFormat=51, ConflictResolution=2)
excel.Application.Quit()
