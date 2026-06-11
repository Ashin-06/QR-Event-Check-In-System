@echo off
title QR Check-In System Launcher
echo ──────────────────────────────────────────────────────────
echo               QR Event Check-In System
echo ──────────────────────────────────────────────────────────
echo.
echo Starting Flask Socket.IO check-in server...
echo.

:: Start the Python server in a new window
start "QR Check-In Server" cmd /k "python app.py"

:: Wait 2 seconds for server startup
timeout /t 2 /nobreak > nul

echo Opening Check-In Dashboard in default browser...
start http://localhost:5001

echo.
echo 🚀 Dashboard opened successfully!
echo.
echo 📡 To connect other phones/tablets on the same Wi-Fi,
echo    open the Network URL shown in the server window.
echo.
echo ──────────────────────────────────────────────────────────
pause
