@echo off
chcp 65001 > nul
title Pub2Doc Studio Launcher
cd /d "%~dp0"
echo Starting Pub2Doc Studio...
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred. Press any key to exit.
    pause > nul
)
