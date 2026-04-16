@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo   LGU Payroll System - Setup ^& Install
echo ==========================================
echo.

:: Check for Node.js
echo [1/3] Checking for Node.js...
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed!
    echo.
    echo Please download and install Node.js (LTS version) from:
    echo https://nodejs.org/
    echo.
    echo After installing Node.js, please run this file again.
    pause
    exit /b
)
echo [OK] Node.js is installed.

:: Install NPM dependencies
echo.
echo [2/3] Installing system dependencies (node_modules)...
echo This process may take 1-3 minutes depending on your internet speed.
echo Please wait...
echo.

call npm install

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install dependencies. 
    echo Please check your internet connection and try again.
    pause
    exit /b
)
echo.
echo [OK] Dependencies installed successfully.

:: Final Check
echo.
echo [3/3] Finalizing setup...
if not exist "data\" mkdir data
echo [OK] Data directory verified.

echo.
echo ==========================================
echo   SETUP COMPLETE!
echo ==========================================
echo.
echo You can now run the system using: Run-Payroll.bat
echo.
pause
