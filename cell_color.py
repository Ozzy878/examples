"""Add an interior color to cells"""
import win32com.client as win32
from pathlib import Path

output_file_name = "cell_color"
output_folder = "S://GitHub/examples/output_files"
output_file = Path('{}/{}.xlsx'.format(output_folder, output_file_name))

excel = win32.gencache.EnsureDispatch('Excel.Application')

wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

for i in range(1, 21):
    ws.Cells(i, 1).Value = i
    ws.Cells(i, 1).Interior.ColorIndex = i

wb.SaveAs(str(output_file), FileFormat=51, ConflictResolution=2)
excel.Application.Quit()
