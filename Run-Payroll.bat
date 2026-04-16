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

:: Delete old files to prevent reading stale data
if exist server_port.txt del server_port.txt
if exist server_pid.txt del server_pid.txt

:: Start Node server in the background and log output
start /B node app.js > server_log.txt 2>&1

:: Wait for server to start (up to 10 seconds)
set /a counter=0
:waitloop
if exist server_port.txt goto :server_ready
timeout /t 1 /nobreak >nul
set /a counter+=1
if %counter% lss 10 goto :waitloop

:: If we get here, server failed to start
echo Set objArgs = WScript.Arguments > "%temp%\msg.vbs"
echo msgbox objArgs(0), 16, "Payroll System Error" >> "%temp%\msg.vbs"
cscript //nologo "%temp%\msg.vbs" "The Payroll System failed to start. Please check server_log.txt for details."
del "%temp%\msg.vbs"
exit /b

:server_ready
:: Give it a tiny bit more time to ensure the server is fully listening
timeout /t 1 /nobreak >nul

:: Read port and PID
set /p PORT=<server_port.txt
set /p PID=<server_pid.txt

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
