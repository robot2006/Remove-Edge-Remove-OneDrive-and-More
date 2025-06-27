@echo off
setlocal

rem Check for Admin rights
>nul 2>&1 net session
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    rem Relaunch this same batch file with admin rights
    powershell.exe -Command "Start-Process -FilePath '%~s0' -Verb RunAs"
    exit /b
)

rem If we are here, we have admin rights.
rem Launch the PowerShell GUI completely hidden AND in Single-Threaded Apartment mode.
powershell.exe -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Windows_Utility_Toolkit.ps1"

endlocal