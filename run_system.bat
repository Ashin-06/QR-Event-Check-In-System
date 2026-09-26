@echo off
setlocal
chcp 65001 >nul
title QR Check-In System
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" launcher.py %*
    goto finished
)
py -3 --version >nul 2>&1
if not errorlevel 1 (
    py -3 launcher.py %*
    goto finished
)
python --version >nul 2>&1
if not errorlevel 1 (
    python launcher.py %*
    goto finished
)
echo [ERROR] Install Python 3.10 or newer, then reopen this launcher.
echo Download: https://www.python.org/downloads/
if /i not "%~1"=="--check" pause
exit /b 1
:finished
set "launcher_result=%errorlevel%"
if not "%launcher_result%"=="0" if /i not "%~1"=="--check" pause
exit /b %launcher_result%
