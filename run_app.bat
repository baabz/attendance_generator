@echo off
setlocal enabledelayedexpansion
title Exam Hall Allocation and Attendance Generator - Kaduna Polytechnic

cd /d "%~dp0"

echo =====================================================================
echo    EXAM HALL ALLOCATION ^& ATTENDANCE GENERATOR
echo    College of Engineering - Kaduna Polytechnic
echo =====================================================================
echo.

:: Step 1: Detect Python
set "PYTHON_CMD="
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
) else (
    py -3 --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON_CMD=py -3"
    )
)

if "%PYTHON_CMD%"=="" (
    echo [ERROR] Python is not installed or not added to your system PATH.
    echo.
    echo Please install Python 3.10, 3.11, or 3.12 from:
    echo https://www.python.org/downloads/
    echo.
    echo NOTE: Ensure you check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

:: Step 2: Check or create Virtual Environment
if not exist "venv\Scripts\activate.bat" (
    echo [1/3] Virtual environment not found. Setting up new environment...
    echo       Creating isolated Python environment in .\venv ...
    %PYTHON_CMD% -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo       Virtual environment created successfully.
    echo.
    echo [2/3] Installing required packages from requirements.txt...
    echo       (This initial download happens only once on first run)
    call "venv\Scripts\activate.bat"
    python -m pip install --upgrade pip --quiet
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install dependencies.
        echo Please check your internet connection for the first-time setup.
        pause
        exit /b 1
    )
    echo       All dependencies installed successfully!
    echo.
) else (
    call "venv\Scripts\activate.bat"
    python -c "import streamlit" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [!] Dependencies missing in existing venv. Installing requirements...
        python -m pip install -r requirements.txt
    )
)

:: Step 3: Launch Application
echo [3/3] Launching Examination Office Web Application...
echo.
echo Opening browser at http://localhost:8501 ...
echo To stop the application, close this command window.
echo =====================================================================
echo.

".\venv\Scripts\python.exe" -m streamlit run app.py --server.port 8501

pause
