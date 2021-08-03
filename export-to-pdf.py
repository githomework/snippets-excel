import win32com.client
from pywintypes import com_error
import sys
from pathlib import Path
import os

if len(sys.argv) != 3:
    WB_PATH = r'C:\apps2\b108_kronos_excel\output\worked_hours_2021-08-02-1000.xlsx'
    PATH_TO_PDF = r'C:\apps2\b108_kronos_excel\output\pdf.pdf'
else:    
    WB_PATH =  sys.argv[1]
    PATH_TO_PDF = sys.argv[2]
f = open(str(Path().absolute()) + "/pdferror.txt", "a")
f.write("starting")

# PDF path when saving
excel = win32com.client.Dispatch("Excel.Application")

excel.Visible = False

try:
    print('Start conversion to PDF')

    # Open
    wb = excel.Workbooks.Open(WB_PATH)

    # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
    ws_index_list = [1]
    wb.WorkSheets(ws_index_list).Select()

    # Save
    wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
except com_error as e:
    print('failed.')
    f.write(e)
else:
    print('Succeeded.')
    f.write("OK")

finally:
    wb.Close()
    excel.Quit()
    f.close()    
    os._exit(0)    
#### Quitting of Excel seems unreliable. May need to kill excel.exe outside of this script.
#### Should make sure Excel is not already running, before running script.
