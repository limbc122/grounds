def findWindows():
    # To figure out window title
    import win32gui
    import pywinauto as pw

    callback = lambda x, y: y.append(x)
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    i = 0
    for hwnd in hwnds:
        title = win32gui.GetWindowText(hwnd)
        if "MSCTFIME" in title or "Default IME" in title:   # PClient
            continue

        if title.strip() == "":
            continue
        
        i += 1
        print("{} : {}".format(str(i), title))

# PC off
# hnd_details = win32gui.GetForegroundWindow()
# print(win32gui.GetWindowText(hnd_details))
# app_details = pw.application.Application(backend="uia").connect(handle=hnd_details)
# 0x0001004C
# 0x5000000E 
# win_details = app_details.window(handle=hnd_details)
# win_details.child_window(title="취소", auto_id="1002", control_type="Button").click_input()
# print(win_details.print_control_identifiers())
# quit()

def findProcesses():
    # To figure out process list
    import win32gui, win32api, win32con, win32process

    i = 0
    for pid in win32process.EnumProcesses():
        i += 1
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ | win32con.PROCESS_SET_INFORMATION , False, pid)
            if handle is None:
                continue

            module_name = win32process.GetModuleFileNameEx(handle, None)

            print("# {} : {}".format(i, module_name))
            if 'NetMan' in module_name:
                print("===============================================================================================>")

        except Exception as e:
            # print("!! Exception : " + str(e))
            continue

if __name__ == "__main__":
    # findWindows()
    findProcesses()