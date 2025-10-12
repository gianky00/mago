@echo off
setlocal

:: ============================================================================
::  Configuration
:: ============================================================================
set "VENV_NAME=venv"
set "PYTHON_SCRIPT=gui.py"

:: ============================================================================
::  Activate virtual environment and run the script
:: ============================================================================
:: Activate the virtual environment
call "%~dp0%VENV_NAME%\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo ERROR: Could not find or activate the virtual environment.
    echo Please run setup.bat first.
    pause
    exit /b 1
)

:: Run the Python script
start "" pythonw.exe "%PYTHON_SCRIPT%"
if %errorlevel% neq 0 (
    echo ERROR: The Python script failed to start.
    pause
    exit /b 1
)

exit /b 0