# About

Simple script that leverages Dell Display Manager to switch display inputs if a USB device is connected or disconnected. Poor man's KVM :)

# Usage
Compile with pyinstaller and install as a service that runs as your user (or the user that DDM is running as)

```commandline
pyinstaller --onefile --hidden-import=win32timezone .\usb_DDM_monitor_service.py --clean --uac-admin
dist\usb_DDM_monitor_service.exe install
dist\usb_DDM_monitor_service.exe start
```
