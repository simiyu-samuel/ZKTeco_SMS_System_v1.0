@echo off
TITLE Biometric Polling Services Stopper

ECHO.
ECHO ==================================================
ECHO   Stopping All Biometric Polling Services...
ECHO ==================================================
ECHO.

ECHO Stopping service for Device A...
taskkill /F /FI "WINDOWTITLE eq Device A Poller"

ECHO Stopping service for Non-Teaching Staff...
taskkill /F /FI "WINDOWTITLE eq Non-Teaching Poller"

REM Add more 'taskkill' commands here for Device B, C, D etc.

ECHO.
ECHO All services have been sent the shutdown command.
ECHO.
pause