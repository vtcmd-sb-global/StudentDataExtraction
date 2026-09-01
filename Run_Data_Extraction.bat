@echo off
title Data Extraction
cd /d "%~dp0"

echo.
echo ========================================
echo   Data Extraction (GUI)
echo ========================================
echo.
echo Starting the graphical interface...
echo.

python data_extraction.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo   Something went wrong.
    echo   Make sure Python and openpyxl are installed:
    echo     pip install openpyxl
    echo ========================================
    pause
)
