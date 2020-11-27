import win32com.client
import win32gui
from time import sleep
import datetime as dt
import os.path
import tkinter as tk
from tkinter import simpledialog, messagebox

print("This program would make you happy very much.")
print("Copyright 2019. bc. All rights reserved.")

shell = win32com.client.Dispatch("Shell.Application")
wscript = win32com.client.Dispatch("WScript.Shell")

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
    messagebox.showwarning("Date", "Input valid date")
    quit()

print("# Day : " + today)

# Get the last internet explorer handler
iexplorer = None
for win in shell.windows():
    if win.Name == "Windows Internet Explorer" or win.Name == "Internet Explorer":
        iexplorer = win

# win32gui.SetForegroundWindow(iexplorer)

if iexplorer is None:
    messagebox.showwarning("Application", "!! Open Internet Explorer from KOSCOM Messenger.")
    print("!! GW should be opened.")
    quit()

# Go to board page
homeURL = "http://148.0.128.21/ekp/main/home/homGwMain"
boardURL = "http://148.0.128.21/ekp/scr/board/boardMain#"
print("# Go to : " + boardURL)
iexplorer.Navigate(homeURL)
while iexplorer.ReadyState != 4:
    sleep(1)
    
iexplorer.Navigate(boardURL)
while iexplorer.ReadyState != 4:
    sleep(1)

doc = iexplorer.Document
while doc.readyState != "complete":
    sleep(1)

try:
    # Click board menu (부서게시판 > 자본시장본부 > 본부게시판 > 일별)
    print("# Select board menu : 부서게시판 > 자본시장본부 > 본부게시판 > 일별")
    sleep(1)
    doc.Body.getElementsByClassName("dynatree-title")[18].click()
    sleep(1)
    doc = iexplorer.Document
    doc.Body.getElementsByClassName("dynatree-title")[19].click()
    sleep(1)
    doc = iexplorer.Document
    doc.Body.getElementsByClassName("dynatree-title")[20].click()
    sleep(1)
    doc = iexplorer.Document
    doc.Body.getElementsByClassName("dynatree-title")[22].click()
    sleep(1)

    # Click 글쓰기 button
    print("# Click the write button")
    doc.Body.getElementsByClassName("btn btn_ico _writeBtn")[0].click()
    while doc.readyState != "complete":
        sleep(1)
    sleep(1)

    # Input title
    week = ('월', '화', '수', '목', '금', '토', '일')
    now = dt.datetime.strptime(today, "%Y%m%d")
    title = "{0}월 {1}일({2}) EXTURE+ 일일종합운영일지".format(str(now.month), str(now.day), week[now.weekday()])
    print("# Input title : " + title)
    doc.forms[0].elements[6].focus()
    doc.forms[0].elements[6].setAttribute('value', title)

    # Input contents
    print("# Input contents")
    content_doc = doc.Body.getElementsByTagName("iframe")[2].contentDocument
    while content_doc.readyState != "complete":
        sleep(1)

    content_doc.Body.getElementsByTagName("iframe")[1].contentDocument.Body.focus()
    sleep(1)
except Exception as e:
    print("!! Exception has occured !")
    messagebox.showerror("Error", str(e) + "\nPlease check GW web page : " + boardURL)
    raise

# Copy excel contents
excel = None
try:
    file_prefix = "EXTURE+일일종합운영일지_{0}".format(today)
    file_dst = os.path.expanduser("~/") + "Documents/{0}.xlsm".format(file_prefix)
    assert os.path.exists(file_dst) == True, "# File dosen't exist. : {0}".format(file_dst)
    print("-> Copy from excel : " + file_dst)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(file_dst)
    wb.Worksheets("일일종합운영일지").Range("A19:D51").Copy()

    # Paste excel contents
    print("-> Paste to the board")
    wscript.SendKeys("^v")
    sleep(2)
except AssertionError as e:
    print("!! AssertionError has occured !")
    messagebox.showerror("Error", str(e))
    raise
except Exception as e:
    print("!! Exception has occured !")
    messagebox.showerror("Error", str(e))
    raise
finally:
    if excel is not None:
        excel.Application.Quit()

# Attach file
doc.body.getElementsByTagName("input")[57].setAttribute("files", file_dst)

print("# Ready to write")
messagebox.showinfo("Success", "* Ready to write *\nPlease select excel file to attach.")