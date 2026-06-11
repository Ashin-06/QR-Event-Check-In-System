@echo off
title QR Check-In Public Internet Tunnel
echo ──────────────────────────────────────────────────────────
echo         QR Event Check-In Public Internet Tunnel
echo ──────────────────────────────────────────────────────────
echo.
echo This tool allows devices on OTHER networks (cellular data,
echo different Wi-Fi, other cities/locations) to connect.
echo.
echo Make sure the Flask server is already running (use run_system.bat).
echo.
echo Establishing secure public tunnel via localhost.run...
echo.
echo ⚠️  INSTRUCTIONS:
echo    1. Look below for a URL ending in '.lhrtunnel.link' (e.g., https://abc.lhrtunnel.link).
echo    2. Open that URL on any phone or tablet on cellular data.
echo    3. Keep this window open. Closing it will terminate the tunnel.
echo.
echo ──────────────────────────────────────────────────────────
ssh -R 80:localhost:5001 nokey@localhost.run
pause
