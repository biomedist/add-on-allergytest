@echo off
setlocal enabledelayedexpansion

:: =====================================================================
::  Allergy Barcode App - Administrator Launcher (self-elevating)
::  - Auto-requests admin rights via UAC
::  - Uses its own folder (%~dp0), so it works regardless of path/Korean
::  - Starts Flask server (app.py), waits for port, opens Chrome
:: =====================================================================

:: --- 1) Check for admin rights; if not, relaunch elevated via UAC ---
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Requesting administrator privileges...
    set "VBS=%temp%\elevate_allergy.vbs"
    echo Set UAC = CreateObject^("Shell.Application"^) > "!VBS!"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "!VBS!"
    "!VBS!"
    del "!VBS!" >nul 2>&1
    exit /B
)

:: --- 2) Now running as admin. Work from this script's own folder ---
pushd "%~dp0"

set "PYTHON=C:\Users\admin\AppData\Local\Programs\Python\Python311\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set PORT=5000
set "CHROME_PROFILE=Profile 22"

:: --- 3) Always stop any running server so the LATEST code is loaded ---
::      (Flask debug reloader does not reliably pick up edits here, so we
::       force a clean restart every launch to avoid running stale code.)
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT% " ^| findstr /I "LISTENING"') do (
    echo Stopping old server PID %%a to load the latest code...
    taskkill /PID %%a /F >nul 2>&1
)
:: give the OS a moment to release the port
timeout /t 1 /nobreak >nul

:: --- 4) Start the server (stderr captured to server_error.txt) ---
start "Allergy Server" /D "%~dp0" cmd /c "%PYTHON% app.py 2> server_error.txt"

:: --- 6) Wait up to 15s for the port to come up ---
set /a COUNT=0
:WAIT
timeout /t 1 /nobreak >nul
netstat -ano | findstr ":%PORT% " | findstr /I "LISTENING" >nul 2>&1
if %errorlevel% == 0 goto OPEN_CHROME
set /a COUNT+=1
if %COUNT% LSS 15 goto WAIT

echo.
echo *** Server failed to start. Check server_error.txt in this folder. ***
if exist "%~dp0server_error.txt" start "" notepad "%~dp0server_error.txt"
pause
exit /B

:OPEN_CHROME
start "" chrome --profile-directory="%CHROME_PROFILE%" "http://localhost:%PORT%"
exit /B
