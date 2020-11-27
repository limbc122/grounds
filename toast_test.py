from win10toast import ToastNotifier

icon_path = "C:/Users/KOSCOM/Documents/workspace/ground/icon/crashdump.ico"

n = ToastNotifier()

n.show_toast("*! Alert !*", "Scheduled job is working", duration=10, icon_path=icon_path) 
