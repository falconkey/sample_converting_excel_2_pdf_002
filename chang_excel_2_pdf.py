import win32com.client
import os
from pathlib import Path, PureWindowsPath
import tkinter as tk
from tkinter import filedialog
from pywintypes import com_error

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilenames()
for single_fullname in file_path:
    print(single_fullname)

    count_total=len(single_fullname)
    real_total=count_total
    print(count_total)
    count_total = count_total - 1

    while count_total > 0:
        if single_fullname[count_total] == "/":
            break
        count_total -= count_total


    print(single_fullname[count_total])


# Path to original excel file
# WB_PATH = r'C:\try_python\vvv.xlsm'
    WB_PATH = r'%s' % single_fullname
#print(file_path)
    ext = '.'+ os.path.realpath(single_fullname).split('.')[-1:][0]
    filefinal = single_fullname.replace(ext,'')
    print(filefinal)
    filefinal = single_fullname + '.pdf'
    newfilefinal = PureWindowsPath(filefinal)
    PATH_TO_PDF = r'%s' % newfilefinal

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        print('Conversing to PDF...')

    # Open
        wb = excel.Workbooks.Open(WB_PATH)

    # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
        ws_index_list = [1]
        wb.WorkSheets(ws_index_list).Select()

    # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except com_error as e:
        print('Conversion Failed.')
    else:
        print('Conversion Succeeded.')
    finally:
        wb.Close()
        excel.Quit()