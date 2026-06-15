@echo off
chcp 65001 > nul
title QR Check-In System Launcher
color 0A

echo ==========================================================
echo               QR Event Check-In System
echo ==========================================================
echo.

:: Check Python is available
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not found. Please install Python 3.8+ and try again.
    echo         Download from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [*] Python found. Starting the server...
echo.

:: Change to the script's directory so relative paths work correctly
cd /d "%~dp0"

:: Launch Flask server in a new persistent window
:: /k keeps the window open even if Python exits, so you can see any error messages
start "QR Check-In Server" cmd /k "chcp 65001 > nul && echo [Server] Starting QR Check-In Server... && echo. && python app.py & echo. & echo [Server] Exited with code %ERRORLEVEL%. Check above for errors. & echo Press any key to close this window. & pause"

:: Wait for server to start
echo [*] Waiting for server to start...
timeout /t 3 /nobreak > nul

echo [*] Opening Check-In Dashboard in browser...
start http://localhost:5001/dashboard

echo.
echo ==========================================================
echo   Server window: "QR Check-In Server" (check for errors)
echo   Dashboard:     http://localhost:5001/dashboard
echo   Scanner URL:   Shown in the server window (Network URL)
echo ==========================================================
echo.
echo   This launcher window can be closed.
echo   The server keeps running in the "QR Check-In Server" window.
echo.
pause
