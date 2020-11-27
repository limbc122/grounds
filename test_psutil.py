import psutil

for proc in psutil.process_iter():
    if "PowerController" in proc.name():
        print("# Name : {}, PID : {}".format(proc.name(), proc.pid))

        process = psutil.Process(proc.pid)
        for child in process.children(recursive=True):
            child.kill()
            print("# Kill Child : {}-{}".format(child.name(), child.pid))
        process.kill()