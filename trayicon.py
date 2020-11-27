import os, logging, threading
from PIL import Image, ImageDraw
from pystray import Icon, Menu, MenuItem
import time, datetime as dt
import win32api, win32con
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR
import eventhook_test as eventhook

print("<Copyright 2019. bc. All rights reserved.>")

# Set logger
logging.basicConfig(format="[%(asctime)s]{%(filename)s:%(lineno)d}-%(levelname)s - %(message)s")
logger = logging.getLogger("My Logger")
logger.setLevel(logging.INFO)

# Set default path
workspace = "C:/Users/KOSCOM/Documents/workspace/"
icon_path = workspace + "ground/icon/Python-icon2.png"
exe_file = workspace + "ground/handle_systray.py"
exe_shift_push = workspace + "/google/mkt_shift_share.py"
exe_on_cal_sync = workspace + "/google/ordernego_sync_dept.py"

# Set scheduler
scheduler = None

def preventScreenSaver():
    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)

def myListener(event):
    if event.exception:
        logger.info("=> The scheduled job crashed !!")
    else:
        if event.job_id not in ["NoScreenSaver", "PrivacyI"]:
            logger.info("=> The scheduled job has worked : {}".format(event.job_id))

def callback(icon):
    global scheduler

    # Create Scheduler
    scheduler = BackgroundScheduler()
    scheduler.add_listener(myListener, EVENT_JOB_EXECUTED | EVENT_JOB_ERROR)
    scheduler.start()

    try:
        # From 06:00 to 22:00
        now = dt.datetime.now()
        logger.info("# Now : {}".format(now))
        # for hour in range(6, 24, 2):
        #     hour = str(hour).zfill(2)
        #     exec_time = dt.datetime.strptime("{}-{}-{} {}:01:00".format(now.year, now.month, now.day, hour), '%Y-%m-%d %H:%M:%S')

        #     # Register only future time
        #     if now < exec_time:
        #         logger.info("# Registered : {}".format(str(exec_time)))
        #         scheduler.add_job(func, 'cron', hour=hour, minute="00", second='30', id="Every2Hour{}".format(hour))

        # Privacy-i and BlackMagic
        # scheduler.add_job(lambda: os.system(exe_file), 'interval', seconds=30, id="PrivacyI")

        # AhnLab V3 : 11:30:30
        # exec_time = dt.datetime.strptime("{}-{}-{} 11:30:30".format(now.year, now.month, now.day), '%Y-%m-%d %H:%M:%S')
        # if now < exec_time:
        #     logger.info("# Registered : {}".format(str(exec_time)))
        #     scheduler.add_job(lambda: os.system(exe_file + " v3"), 'cron', hour="11", minute="30", second='30', id="ForV3")

        # No Screen Saver
        logger.info("# Registered : No Screen Saver")
        scheduler.add_job(preventScreenSaver, 'interval', minutes=4, id="NoScreenSaver")

        # Telegram Push
        logger.info("# Registered : Telegram Shift Push")
        scheduler.add_job(lambda: os.system(exe_shift_push), 'cron', day_of_week="0-4", hour="16", minute="05", second='01', id="TelegramShiftPush")

        # ON Calendar Sync
        logger.info("# Registered : ON Calendar Sync")
        scheduler.add_job(lambda: os.system(exe_on_cal_sync), 'interval', hours=1, id="ONCalendarSync")

        logger.info("## Adding jobs success !")
    except Exception as e:
        scheduler.shutdown()
        logger.exception(str(e))

    icon.visible = True


menu_items = [
    # MenuItem('Force Execution', lambda: os.system(exe_file)),
    MenuItem('Telegram Shift Push', lambda: os.system(exe_shift_push)),
    MenuItem('ON Calendar Sync', lambda: os.system(exe_on_cal_sync)),
    MenuItem('Show Scheduled Jobs', lambda: scheduler.print_jobs()),
    # MenuItem('AhnLab V3', on_clicked, checked=lambda item: state),
    MenuItem("Exit", lambda: icon.stop()),
]
menu = Menu(*menu_items)

icon = Icon("Kill Shits", menu=menu)
icon.title = "Kill shits"
icon.icon = Image.open(icon_path)

# Set Event Driven Method
t = threading.Thread(target=eventhook.run)
t.daemon = True
t.start()

# Icon run
logger.info("## Icon running")
icon.run(setup=callback)

logger.info("## Quit")
