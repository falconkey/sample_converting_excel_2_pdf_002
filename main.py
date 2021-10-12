import win32com.client
import tkinter as tk
from tkinter import filedialog
from pywintypes import com_error

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
print(file_path)

# Path to original excel file
# WB_PATH = r'C:\try_python\vvv.xlsm'
WB_PATH = file_path

# PDF path when saving
# PATH_TO_PDF = r'C:\try_python\vvv.pdf'
PATH_TO_PDF = r'C:\try_python\vvvvvv.pdf'

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