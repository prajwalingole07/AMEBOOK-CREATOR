@echo off
title AME CUTOFF BOOK MAKER - Admissions Made Easy
color 0B
cd /d "%~dp0"

echo ============================================================
echo     STARTING AME CUTOFF BOOK MAKER LOCALHOST WEB APP
echo     Admissions Made Easy, Latur - Contact: 94226 11661
echo ============================================================
echo.
echo Server running at: http://localhost:5000
echo Opening Chrome...
echo.

:: Open browser after 2 seconds invisibly without creating a second CMD window
start "" powershell -windowstyle hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:5000'"

:: Run Flask server directly in this single CMD window
python app.py
