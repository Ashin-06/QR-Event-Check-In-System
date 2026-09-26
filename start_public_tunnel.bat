@echo off
setlocal
title QR Check-In Public Tunnel
call "%~dp0run_system.bat" --tunnel %*
exit /b %errorlevel%
