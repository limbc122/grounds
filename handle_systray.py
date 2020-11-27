import warnings
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.simplefilter("ignore", category=FutureWarning)

import win32gui, win32api, win32con, win32process
import pywinauto as pw
from pywinauto import taskbar
import time, logging
from win10toast import ToastNotifier
import sys
# from ctypes import windll

print("<Copyright 2019. bc. All rights reserved.>")

# Set logger
logging.basicConfig(format="[%(asctime)s]{%(filename)s:%(lineno)d}-%(levelname)s - %(message)s")
logger = logging.getLogger("My Logger")
logger.setLevel(logging.INFO)

# PC off
# hnd_details = win32gui.GetForegroundWindow()
# print(win32gui.GetWindowText(hnd_details))
# app_details = pw.application.Application(backend="uia").connect(handle=hnd_details)
# win_details = app_details.window(handle=hnd_details)
# win_details.child_window(title="취소", auto_id="1002", control_type="Button").click_input()
# print(win_details.print_control_identifiers())
# quit()

logger.info("## Start handling processes ##")

is_PIAutoScan = False

# Get Processes
logger.info("# Get processes")
i = 0 
for pid in win32process.EnumProcesses():
    i += 1
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        if handle is None:
            continue

        module_name = win32process.GetModuleFileNameEx(handle, None)

        if "PIAutoScan" in module_name:
            # is_PIAutoScan = True
            path_PIAutoScan = module_name
            logger.info("! PIAutoScan is working : {}".format(path_PIAutoScan))

        # print("# {} : {}".format(i, module_name))
        if "BlackMagic" in module_name:
            logger.info("! BlackMagic is working: {}".format(module_name))
            # win32api.TerminateProcess(handle,-1)
            # win32api.CloseHandle(handle)
            # logger.info("! BlackMagic is terminated.")

    except Exception as e:
        # print("!! Exception : " + str(e))
        continue

if not (is_PIAutoScan or (len(sys.argv) > 1 and sys.argv[1] == "v3")):
    quit()

# Toast notification
icon_path = "D:/workspace/ground/icon/warning.ico"
toast = ToastNotifier()
toast.show_toast("*! Alert !*", "=> Killing shit processes get started.", duration=3, icon_path=icon_path, threaded=False)

# Get Shell Tray Handle
logger.info("# Get shell tray handle")
hnd_taskbar = pw.findwindows.find_elements(class_name="Shell_TrayWnd")[0].handle
app_taskbar = pw.application.Application().connect(handle=hnd_taskbar)
win_taskbar = app_taskbar.ShellTrayWnd.TrayNotifyWnd
hnd_button_showicons = win_taskbar.child_window(class_name="Button", found_index=0)

# Freeze mouse and keyboard input
# windll.user32.BlockInput(True)

# Get hidden system tray icons
logger.info("# Get system tray")
logger.info("# Show hidden tray icon window")
app_taskbar.ShellTrayWnd.click_input()
hnd_button_showicons.click_input()
dlg_systray = taskbar.explorer_app.window(class_name='NotifyIconOverflowWindow')
popup_toolbar = dlg_systray.child_window(class_name="ToolbarWindow32")
# print(popup_toolbar.texts()[1:])

button_privacyi = None
button_v3 = None
for txt in popup_toolbar.texts():
    if "Privacy-i" in txt or "진행률" in txt:
        button_privacyi = popup_toolbar.button(txt)
    elif "AhnLab V3 Internet Security" in txt:
        button_v3 = popup_toolbar.button(txt)


if len(sys.argv) > 1 and sys.argv[1] == "v3":
    logger.info("# Check whether V3 is working")
    button_v3.click_input(button='right', double=False)

    cnt = 0
    popup_menu = None
    while popup_menu is None and cnt <= 30:
        try:
            popup_menu = pw.application.Application(backend="uia").connect(class_name="#32768")
            popup_menu = popup_menu.window(class_name="#32768")
        except:
            time.sleep(1)
            cnt += 1
            continue

    if popup_menu is None:
        logger.info("!! Error occured during getting popup menu window !!")
        quit()

    is_V3Scan = False
    for item in popup_menu.wrapper_object().items():
        if item.texts()[0] == "예약 검사 중지(D)":
            logger.info("# Quit V3 working")
            is_V3Scan = True
            item.select()
            pw.keyboard.send_keys('{ENTER}')

    if is_V3Scan == False:
        logger.info("# V3 is not working")

    app_taskbar.ShellTrayWnd.click_input()
    hnd_button_showicons.click_input()

if is_PIAutoScan == True:
    logger.info("# Privacy-i is working")
    button_privacyi.click_input(button='right', double=False)

    # Select the Menu
    logger.info("# Select Menu")
    cnt = 0
    popup_menu = None
    while popup_menu is None and cnt <= 30:
        try:
            popup_menu = pw.application.Application(backend="uia").connect(class_name="#32768")
            popup_menu = popup_menu.window(class_name="#32768")
        except:
            time.sleep(1)
            cnt += 1
            continue    
    
    if popup_menu is None:
        logger.info("!! Error occured during getting popup menu window !!")
        quit()
    else:
        popup_menu.wrapper_object().item_by_index(0).select()

    logger.info("# Get popup window from system tray")
    app_PI_dialog = None
    while app_PI_dialog is None:    
        try:
            app_PI_dialog = pw.application.Application(backend="uia").connect(class_name="PIClientMainDialog")
            win_PI_dialog = app_PI_dialog.window(class_name="PIClientMainDialog")
        except:
            time.sleep(1)
            continue
            
    win_PI_dialog.wait('visible', timeout=20)

    logger.info("# Click 상세 보기")
    win_PI_dialog.child_window(title="상세 보기").click_input()

    hnd_details = win32gui.GetForegroundWindow()
    app_details = pw.application.Application(backend="uia").connect(handle=hnd_details)
    win_details = app_details.window(handle=hnd_details)
    
    logger.info("# Click 중지")
    win_details.child_window(title="중지").click_input()

    logger.info("# Kill popup screen")
    hnd_details = win32gui.GetForegroundWindow()
    app_details = pw.application.Application(backend="uia").connect(handle=hnd_details)
    win_details = app_details.window(handle=hnd_details)

    win_details.child_window(title="닫기", found_index=0).click_input()
    win_PI_dialog.child_window(title="닫기", found_index=0).click_input()
    # app_details.kill()
else:
    logger.info("# Privacy-i is not working")
    hnd_button_showicons.wrapper_object().click_input()

logger.info("# Process management finished")

toast.show_toast("*! Alert !*", "=> Killing shit processes has finished.", duration=2, icon_path=icon_path, threaded=False)