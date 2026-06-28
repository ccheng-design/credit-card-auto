@echo off
cd /d "%~dp0"

python "credit-card-auto.py"

if errorlevel 1 (
    echo.
    echo The script exited with an error.
    pause
)