@echo off
title Windows Deep Cleaner Launcher
echo ========================================
echo Starting Windows Deep Cleaner...
echo Requesting Admin rights and bypassing policy...
echo ========================================
echo.
powershell -Command "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File "%~dp0Win-Deep-Cleaner.ps1"' -Verb RunAs"
pause