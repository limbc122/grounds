import win32com.client
import win32api
import datetime as dt
import os.path
import glob
import tkinter as tk
from tkinter import simpledialog, messagebox

print("This program would make you happy very much.")
print("Copyright 2019. bc. All rights reserved")

# Input date
root = tk.Tk()
root.withdraw()
today = dt.datetime.now().strftime("%Y%m%d")
today = simpledialog.askstring("Date", "Input Date (YYYYMMDD)", parent=root, initialvalue=today)
if today is None:
    messagebox.showwarning("Date", "Input valid date")
    quit()

try:
    dt.datetime.strptime(today, "%Y%m%d")
except ValueError:
    print("!! Wrong date format.")
    messagebox.showwarning("Date", "Wrong date format")
    quit()

print("# Day : " + today)

try:
    # Set file names
    file_prefix = "EXTURE+일일종합운영일지_{0}".format(today)
    folder_name = "/★★Exture+ 일일종합운영일지/"

    # Find out base path    
    for drive in win32api.GetLogicalDriveStrings().split('\\\000'):
        base_path = drive + folder_name
        if os.path.exists(base_path) == True:
            print("# Base path : " + base_path)
            break

    # Make target path
    file_match = base_path + "부문별작성폴더/{0}_현선물파생.xlsm".format(file_prefix)
    file_tech = base_path + "부문별작성폴더/{0}_기반기술.xlsx".format(file_prefix)
    file_infra = base_path + "부문별작성폴더/{0}_인프라.xlsb".format(file_prefix)
    # file_dst = base_path + "{0}.xlsm".format(file_prefix)
    file_dst = os.path.expanduser("~/") + "Documents/{0}.xlsm".format(file_prefix)

    # Find the latest written file
    file_srcs = glob.glob(base_path + "EXTURE*.xlsm")
    file_srcs.sort()
    file_src = base_path + os.path.basename(file_srcs[-1])
except Exception as e:
    print("!! Exception has occured !")
    messagebox.showerror("Error", str(e) + "\nPlease check network drive : " + base_path)
    raise

excel = None
try:
    # Check file path
    assert os.path.exists(file_src) == True, "# File dosen't exist. : {0}".format(file_src)
    assert os.path.exists(file_match) == True, "# File dosen't exist. : {0}".format(file_match)
    assert os.path.exists(file_tech) == True, "# File dosen't exist. : {0}".format(file_tech)

    # Open excel application
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Open excel file
    print("# Open source excel file : {0}".format(file_src))
    wb_src = excel.Workbooks.Open(file_src)

    # Copy and Paste : Matching
    print("# Copy and Paste : Matching")

    wb_match = excel.Workbooks.Open(file_match)
    # wb_match.Worksheets("EXTURE+ 현선물 통계").UsedRange.Copy()
    ws_match_sec = wb_match.Worksheets("EXTURE+ 현선물 통계")
    ws_match_sec.Range(ws_match_sec.Cells(1,1), ws_match_sec.Cells(ws_match_sec.UsedRange.Rows.Count, ws_match_sec.UsedRange.Columns.Count)).Copy()
    ws = wb_src.Worksheets("EXTURE+ 현선물 통계")
    ws.Paste(ws.Cells(1,1))

    # wb_match.Worksheets("EXTURE+ 채권 통계").UsedRange.Copy()
    ws_match_bnd = wb_match.Worksheets("EXTURE+ 채권 통계")
    ws_match_bnd.Range(ws_match_bnd.Cells(1,1), ws_match_bnd.Cells(ws_match_bnd.UsedRange.Rows.Count, ws_match_bnd.UsedRange.Columns.Count)).Copy()
    ws = wb_src.Worksheets("EXTURE+ 채권 통계")
    ws.Paste(ws.Cells(1,1))

    wb_match.Close()

    # Copy and Paste : Infra
    print("# Copy and Paste : Infra")

    wb_infra = excel.Workbooks.Open(file_infra)
    wb_infra.Worksheets("일간시스템사용률").UsedRange.Copy()
    ws = wb_src.Worksheets("일간시스템사용률")
    ws.Paste(ws.Cells(1,1))
    wb_infra.Close()

    # Copy and Paste : Tech
    print("# Copy and Paste : Tech")
    wb_tech = excel.Workbooks.Open(file_tech)

    # Delete for update
    ws = wb_src.Worksheets("RT,TAT 현황")

    ws_tech = wb_tech.Worksheets(1)
    ws_tech.Range(ws_tech.Cells(4,1), ws_tech.Cells(ws_tech.UsedRange.Rows.Count, ws_tech.UsedRange.Columns.Count)).Copy()
    ws.Paste(ws.Cells(2,1))
    if ws.UsedRange.Rows.Count >= 2+(ws_tech.UsedRange.Rows.Count-4)+1:
        ws.Range(ws.Cells(2+(ws_tech.UsedRange.Rows.Count-4)+1,1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).Delete()

    wb_tech.Close()
    
    # Save excel as a new file
    print("# Save as a new file : {0}".format(file_dst))
    wb_src.Worksheets(1).Activate()
    wb_src.SaveCopyAs(Filename=file_dst)
except AssertionError as e:
    print("!! AssertionError has occured !")
    messagebox.showerror("Error", str(e))
    raise
except Exception as e:
    print("!! Exception has occured !")
    messagebox.showerror("Error", str(e))
    raise
finally:
    # Close excel application
    if excel is not None:
        excel.Application.Quit()
        print("# Excel Quit")

print("## Making an integrated excel file is succeeded.")
messagebox.showinfo("Success", "* Please check the 'my documents' folder.")