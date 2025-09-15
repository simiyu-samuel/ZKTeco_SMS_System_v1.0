@echo off
TITLE Biometric Polling Services Control Panel

ECHO.
ECHO ===============================================
ECHO   Starting All Biometric Polling Services...
ECHO ===============================================
ECHO.

REM IMPORTANT: Update these paths to match your folder structure!

ECHO Starting service for Device A...
start "Device A Poller" python C:\ZKTeco_SMS_System_v1.0\Device_Templates\DeviceA\zkteco.py

ECHO Starting service for Non-Teaching Staff...
start "Non-Teaching Poller" python C:\ZKTeco_SMS_System_v1.0\Device_Templates\NonTeaching\zkteco.py

REM Add more 'start' commands here for Device B, C, D etc.

ECHO.
ECHO All services have been launched in separate windows.
ECHO You can close this control panel window now.
ECHO.
pause