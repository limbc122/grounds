import warnings
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.simplefilter("ignore", category=FutureWarning)

import win32gui, win32api, win32con
import pywinauto as pw
from pywinauto.application import Application
from pywinauto import taskbar
import time, logging

# Set logger
logging.basicConfig(format="[%(asctime)s]{%(filename)s:%(lineno)d}-%(levelname)s - %(message)s")
logger = logging.getLogger("My Logger")
logger.setLevel(logging.INFO)

def mouse_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

# mouse_click(1018, 647)
# pw.mouse.click(button='left', coords=(1018, 647))

# win_list = pw.findwindows.enum_windows()
# print(win_list)
# wins = pw.findwindows.find_windows()
# for win in wins:
#     print(win32gui.GetWindowText(win))

# for i, win in enumerate(wins):
#     if  i == 7:
#         print(f'[{i}] ' + win32gui.GetWindowText(win))
#         app = pw.application.Application(backend="uia").connect(handle=win)
#         app_win = app.window(handle=win)
#         print(app_win.print_control_identifiers(depth=1))

# Cancel PC off
# 1) Find window handle
# hnd_pcoff = pw.findwindows.find_elements(class_name="Shell_TrayWnd")[0].handle
# hnd_pcoff = win32gui.GetForegroundWindow()
# print(win32gui.GetWindowText(hnd_pcoff))

# 2) Get application from window handle
# app = Application().connect(title_re=".*Notepad", class_name="Notepad")
# app_pcoff = pw.application.Application(backend="uia").connect(handle=hnd_pcoff)
# win_pcoff = app_pcoff.window(handle=hnd_pcoff)
# dlg = app.top_window()
# win_pcoff.child_window(title="취소", auto_id="1002", control_type="Button").click_input()
# print(win_details.print_control_identifiers())
# quit()

# app.YourDialog.print_control_identifiers()


import os
os.system("wmic process where \"name='PowerController.exe'\" delete")