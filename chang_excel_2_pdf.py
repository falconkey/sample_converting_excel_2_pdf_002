import win32com.client
import os
from pathlib import Path, PureWindowsPath
import tkinter as tk
from tkinter import filedialog
from pywintypes import com_error

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilenames()
for single_full_filename in file_path:
    print(single_full_filename)

    no_of_ch_single_full_filename = len(single_full_filename)
    print(no_of_ch_single_full_filename)
    counter_1 = no_of_ch_single_full_filename

    while True:
        if single_full_filename[counter_1 -1] == "/" :
            break
        else:
            counter_1 -= 1

    loc_for_first_char = counter_1 + 5 #counting from last "/" to real name
    #loc_for_last_char = loc_for_first_char + 9
    print(single_full_filename)
    single_full_ture_filename = single_full_filename[loc_for_first_char:]
    single_full_ture_filename = single_full_ture_filename[:9]
    single_full_ture_filename = single_full_ture_filename.lower()

#    print(single_full_ture_filename)


    single_full_ture_filename_path = single_full_filename[:loc_for_first_char-5]

 #   print(single_full_ture_filename_path)

    single_full_ture_filename_with_path = single_full_ture_filename_path + single_full_ture_filename

    print(single_full_ture_filename_with_path)

# Path to original excel file
# WB_PATH = r'C:\try_python\vvv.xlsm'
    WB_PATH = r'%s' % single_full_filename
#print(file_path)
#    ext = '.'+ os.path.realpath(single_full_filename).split('.')[-1:][0]
#    filefinal = single_full_filename.replace(ext,'')
    filefinal = single_full_ture_filename_with_path + '.pdf'
    newfilefinal = PureWindowsPath(filefinal)
    PATH_TO_PDF = r'%s' % newfilefinal

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        print('Conversing to PDF in Progress')

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