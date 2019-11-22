"""payrates.py
Report payrates for two employees across multiple spreadsheets"""
import win32com.client as win32
import glob
import os

xlxsfiles = sorted(glob.glob("Payroll/*.xlsx"))
print("Reading %d files..." % len(xlxsfiles))

austin = []
john_doe = []
cwd = os.getcwd()
excel = win32.gencache.EnsureDispatch('Excel.Application')

fpjohn_doeaustin = open('output_files/austin_johndoe.csv', 'w')

for xlsxfile in xlxsfiles:
    wb = excel.Workbooks.Open(cwd + "\\" + xlsxfile)
    try:
        ws = wb.Sheets('PAYROLL')
    except:
        print("No sheet named 'PAYROLL' in %s, skipping" % xlsxfile)
        john_doe.append(0.0)
        austin.append(0.0)
        wb.Close()

        continue

    xldata = ws.UsedRange.Value
    names = [r[1] for r in xldata]

    if u'WALDRON, AUSTIN' in names:
        indx = names.index(u'WALDRON, AUSTIN')
        austin.append(xldata[indx][4])
    else:
        austin.append(0)

    if u'DOE, JOHN' in names:
        indx = names.index(u'DOE, JOHN')
        john_doe.append(xldata[indx][4])
    else:
        john_doe.append(0)

    wb.Close()

fpjohn_doeaustin.write("File,John,Austin\n")

for i in range(len(xlxsfiles)):
    fpjohn_doeaustin.write("%s,%0.2f,%0.2f\n" % (xlxsfiles[i], john_doe[i], austin[i]))

print("Wrote john_doeaustin.csv")

excel.Application.Quit()
