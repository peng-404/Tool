@echo off
title Win Deep Cleaner Launcher
echo ========================================
echo Starting Win Deep Cleaner...
echo Launching PowerShell script...
echo ========================================
echo.
set "SCRIPT_PATH=%~dp0Win-Deep-Cleaner.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
if errorlevel 1 (
    echo.
    echo Win Deep Cleaner failed to start correctly.
    pause
)
