@echo off
if "%~1"=="hidden" goto :main

echo ==========================================
echo   LGU Payroll Management System Launcher
echo ==========================================
echo.

:: Check for Node.js before hiding
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed!
    echo Please run Install-Dependencies.bat first.
    pause
    exit /b
)

echo Starting system... Please wait.

:: Create VBS to run this batch file hidden
set "vbs=%temp%\run_payroll_hidden.vbs"
echo Set WshShell = CreateObject("WScript.Shell") > "%vbs%"
echo WshShell.Run chr(34) ^& "%~f0" ^& chr(34) ^& " hidden", 0, False >> "%vbs%"
cscript //nologo "%vbs%"
del "%vbs%"
exit /b

:main
cd /d "%~dp0"

:: Start Node server in the background
start /B node app.js

:: Wait for server to start and write the port/PID files
timeout /t 3 /nobreak >nul

:: Read port and PID
set PORT=3000
if exist server_port.txt set /p PORT=<server_port.txt

set PID=
if exist server_pid.txt set /p PID=<server_pid.txt

:: Find Chrome or Edge to open as a standalone app window
set "BROWSER="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "BROWSER=C:\Program Files\Google\Chrome\Application\chrome.exe"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "BROWSER=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" set "BROWSER=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" set "BROWSER=C:\Program Files\Microsoft\Edge\Application\msedge.exe"

if defined BROWSER (
    :: Run the browser in app mode with a separate profile.
    :: This blocks the script until the user closes the window.
    "%BROWSER%" --app="http://127.0.0.1:%PORT%" --user-data-dir="%temp%\PayrollAppProfile"
) else (
    :: Fallback if no supported browser is found
    start http://127.0.0.1:%PORT%
    timeout /t 86400 >nul
)

:: When the window is closed, kill the specific node process
if not "%PID%"=="" (
    taskkill /F /PID %PID% >nul 2>&1
)

exit /b
