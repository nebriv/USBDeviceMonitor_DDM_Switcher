import os
import win32api
import win32file
import win32com.client
import datetime
import time
import wmi
import json
import subprocess
import threading
import logging


# Device Classes (Could be keyboard, mouse, etc run out put from PowerShell -Command \"& {Get-PnpDevice | Select-Object Status,Class,FriendlyName,InstanceId)
monitored_device_classes = ["mouse"]

# Monitored Device Names (Change this)
monitored_device_names = ["Razer DeathAdder Elite"]

# Input names to switch to
disconnected_input = "hdmi2"
connected_input = "dp"

# Location of Dell Display Manager EXE
ddm_exe = "C:\\Program Files\\Dell\\Dell Display Manager 2.0\\DDM.exe"


log = logging.getLogger(__name__)

class Monitor:
    def __init__(self):
        log.info("Init")
        self.run = False
        try:
            out = subprocess.check_output(["%s" % ddm_exe])
            log.info("DDM Output:\n%s" % out)
        except Exception as err:
            log.error("Error running ddm:\n%s" % err)

    def start(self):
        log.info("Starting")
        self.run = True
        self.main()

    def stop(self):
        self.run = False

    def main(self):
        log.info("Monitor running, looking for %s connection(s)..." % (", ".join(monitored_device_names)))
        run_count = 0
        devices = []
        exists = None
        last = None
        first_run = True
        while self.run:
            out = subprocess.getoutput(
                "PowerShell -Command \"& {Get-PnpDevice | Select-Object Status,Class,FriendlyName,InstanceId | ConvertTo-Json}\"")
            j = json.loads(out)

            found_classes = []
            exists = False
            for dev in j:
                if str(dev['Class']).lower() in monitored_device_classes:
                    if str(dev['FriendlyName']) in monitored_device_names:
                        if dev['Status'] == "OK":
                            exists = True
                        # print(dev['Status'], dev['Class'], dev['FriendlyName'], dev['InstanceId'])

            # print("Exists: %s - Last: %s" % (exists, last))
            if first_run:
                first_run = False
                last = exists
            else:
                if exists != last:
                    log.info("Change Detected")
                    print("Change Detected")
                    if exists == True:
                        print("Mouse detected, setting active input to DisplayPort")
                        log.info("Mouse detected, setting active input to DisplayPort")
                        out = subprocess.check_output(["%s" % ddm_exe, "/writeactiveinput", "%s" % connected_input])
                        log.info("Command Output:\n%s" % out)
                        print("Command Output:\n%s" % out)
                    else:
                        print("Mouse not detected, setting active input to HDMI 1")
                        log.info("Mouse not detected, setting active input to HDMI 1")
                        out = subprocess.call(["%s" % ddm_exe, "/writeactiveinput", "%s" % disconnected_input])
                        log.info("Command Output:\n%s" % out)
                        print("Command Output:\n%s" % out)

                    last = exists
        log.info("Out of while loop")

if __name__ == '__main__':
    monitor = Monitor()
    monitor.start()