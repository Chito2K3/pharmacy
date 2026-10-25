@echo off
setlocal
title ARMMC Pharmacy IMS -- Live Google Sheet Sync
cd /d "%~dp0\.."

echo =================================================================
echo   Amang Rodriguez Memorial Medical Center (ARMMC)
echo   Pharmacy Inventory Monitoring System -- Sync Utility
echo =================================================================
echo.
echo Connecting to live Google Sheet and refreshing items ^& reorders...
echo.

python tools\sync_from_google_sheet.py
if %ERRORLEVEL% EQU 0 (
    echo.
    echo =================================================================
    echo   Sync completed successfully! 
    echo   Both items.json and reorders.json have been updated.
    echo =================================================================
) else (
    echo.
    echo [ERROR] Synchronization encountered an issue. Check Python setup.
)

echo.
pause
