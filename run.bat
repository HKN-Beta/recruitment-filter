@echo off
REM HKN Recruitment Filter - Windows Run Script
REM This script installs required Python libraries and runs the recruitment filter

echo ========================================
echo HKN Recruitment Filter Setup and Run
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://python.org and ensure it's added to PATH
    echo.
    pause
    exit /b 1
)

echo Python detected:
python --version
echo.

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: pip is not installed or not available
    echo Please ensure pip is installed with Python
    echo.
    pause
    exit /b 1
)

echo Installing required Python packages...
echo ========================================

REM Install requirements
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install required packages
    echo Please check your internet connection and try again
    echo.
    pause
    exit /b 1
)

echo.
echo Package installation completed successfully!
echo.

echo Running HKN Recruitment Filter...
echo ========================================
echo.

REM Run the main Python script
python hkn_student_parser.py
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Script execution failed
    echo Please check the error messages above
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Script completed successfully!
echo.
echo Generated files:
echo - seniors.csv
echo - juniors.csv  
echo - sophomores.csv
echo - report.txt
echo.
echo Check the report.txt file for detailed statistics and any errors.
echo.
pause
