import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
from monitor import Monitor
import threading
import concurrent.futures
import time
import logging
import os

logging.basicConfig(filename="C:\\temp\\USBMouseMonitor.log", level=logging.DEBUG)

log = logging.getLogger(__name__)

log.info("Top of script - running as %s" % os.getlogin())

class workingthread(threading.Thread):
    def __init__(self, quitEvent):
        self.quitEvent = quitEvent
        self.waitTime = 1
        threading.Thread.__init__(self)

    def run(self):
        log.info("Run")
        try:
            # Running start_flask() function on different thread, so that it doesn't block the code
            executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            executor.submit(self.start_monitor)
        except Exception as err:
            log.error("Got error while running thread\n%s" % err)
            pass

        # Following Lines are written so that, the program doesn't get quit
        # Will Run a Endless While Loop till Stop signal is not received from Windows Service API
        while not self.quitEvent.isSet():  # If stop signal is triggered, exit
            time.sleep(1)

        self.test.stop()

    def start_monitor(self):
        # This Function contains the actual logic, of windows service
        # This is case, we are running our flaskserver
        self.test = Monitor()
        self.test.start()

# https://stackoverflow.com/questions/32404/how-do-you-run-a-python-script-as-a-service-in-windows
class AppServerSvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "USBMonitor_DisplaySwitcher"
    _svc_display_name_ = "USB Monitor Display Switcher"

    def __init__(self, args):
        log.info("Init of service")
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = threading.Event()
        self.thread = workingthread(self.hWaitStop)

    def SvcStop(self):
        log.info("Service stop")
        self.hWaitStop.set()

    def SvcDoRun(self):
        log.info("Service start")
        self.thread.start()
        self.hWaitStop.wait()
        self.thread.join()


if __name__ == '__main__':
    try:
       if len(sys.argv) > 1:
           # Called by Windows shell. Handling arguments such as: Install, Remove, etc.
           win32serviceutil.HandleCommandLine(AppServerSvc)
       else:
           # Called by Windows Service. Initialize the service to communicate with the system operator
           servicemanager.Initialize()
           servicemanager.PrepareToHostSingle(AppServerSvc)
           servicemanager.StartServiceCtrlDispatcher()
    except Exception as err:
        log.error("Error in main:\n%s" % err)