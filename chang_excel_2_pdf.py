import win32com.client
import os
from pathlib import Path, PureWindowsPath
import tkinter as tk
from tkinter import filedialog
from pywintypes import com_error

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
print(file_path)

# Path to original excel file
# WB_PATH = r'C:\try_python\vvv.xlsm'
WB_PATH = r'%s' % file_path
print(file_path)
ext = '.'+ os.path.realpath(file_path).split('.')[-1:][0]
filefinal = file_path.replace(ext,'')
filefinal = file_path + '.pdf'

newfilefinal = PureWindowsPath(filefinal)

PATH_TO_PDF = r'%s' % newfilefinal

# PDF path when saving
#PATH_TO_PDF = r'C:\try_python\vvv.pdf'
print(PATH_TO_PDF)

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
else:
    print('Succeeded.')
finally:
    wb.Close()
    excel.Quit()