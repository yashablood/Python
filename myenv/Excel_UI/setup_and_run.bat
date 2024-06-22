@echo off
REM Check if Python is installed
python --version
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed. Please install Python from https://www.python.org/downloads/
    pause
    exit /b
)

REM Install required libraries
pip install pandas openpyxl tkcalendar matplotlib

REM Run the Python script
python script.py
