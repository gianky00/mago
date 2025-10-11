@echo off
setlocal enabledelayedexpansion

:: ============================================================================
::  Request administrator privileges
:: ============================================================================
echo Requesting administrator privileges...
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Requesting to elevate privileges...
    goto UACPrompt
) else (
    goto gotAdmin
)

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "!~s0!", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

:: ============================================================================
::  Configuration
:: ============================================================================
set "VENV_NAME=venv"
set "PYTHON_SCRIPT=gui.py"
set "REQUIREMENTS_FILE=requirements.txt"

:: ============================================================================
::  Find the Python executable
:: ============================================================================
echo.
echo Searching for a valid Python installation...

set "PYTHON_EXE="
for /f "delims=" %%i in ('where python 2^>nul') do (
    if not defined PYTHON_EXE (
        echo "%%i" | find /I "\WindowsApps\" >nul
        if errorlevel 1 (
            set "PYTHON_EXE=%%~i"
        )
    )
)

if not defined PYTHON_EXE (
    echo Fallback: Searching in common installation directories...
    for /r "%ProgramFiles(x86)%" %%f in (python.exe) do (
        if not defined PYTHON_EXE (
            echo "%%f" | find /I "\WindowsApps\" >nul
            if errorlevel 1 ( set "PYTHON_EXE=%%~f" )
        )
    )
    for /r "%ProgramFiles%" %%f in (python.exe) do (
        if not defined PYTHON_EXE (
            echo "%%f" | find /I "\WindowsApps\" >nul
            if errorlevel 1 ( set "PYTHON_EXE=%%~f" )
        )
    )
)

if not defined PYTHON_EXE (
    echo.
    echo ERROR: Could not find a valid Python installation.
    goto :error
)

echo Found Python at: %PYTHON_EXE%
echo.

:: ============================================================================
::  Create and activate the virtual environment (Robust Method)
:: ============================================================================
if not exist "%~dp0%VENV_NAME%\Scripts\activate.bat" (
    echo Creating virtual environment '%VENV_NAME%'...

    :: Step 1: Create a minimal environment without pip to avoid AV conflicts.
    "%PYTHON_EXE%" -m venv "%~dp0%VENV_NAME%" --without-pip
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create the basic virtual environment structure.
        goto :error
    )

    echo Activating the minimal environment...
    call "%~dp0%VENV_NAME%\Scripts\activate.bat"

    echo Installing/upgrading pip inside the environment...
    :: Step 2: Use ensurepip to robustly install pip in the created venv.
    python.exe -m ensurepip --upgrade
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install pip in the virtual environment.
        goto :error
    )
    echo Virtual environment created and pip is ready.

) else (
    echo Virtual environment '%VENV_NAME%' already exists.
    echo Activating the virtual environment...
    call "%~dp0%VENV_NAME%\Scripts\activate.bat"
)
echo.

:: ============================================================================
::  Install dependencies
:: ============================================================================
echo Installing dependencies from '%REQUIREMENTS_FILE%'...
python.exe -m pip install -r "%REQUIREMENTS_FILE%" --no-cache-dir --upgrade
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    goto :error
)
echo Dependencies installed successfully.
echo.

:: ============================================================================
::  Run the main Python script
:: ============================================================================
echo Starting the script '%PYTHON_SCRIPT%'...
echo.
python.exe "%PYTHON_SCRIPT%"
if %errorlevel% neq 0 (
    echo ERROR: The Python script exited with an error.
    goto :error
)

echo.
echo The script ran successfully.
goto :end

:error
echo.
echo An error occurred. The window will remain open for analysis.
pause >nul

:end
echo.
echo Press any key to close the window.
pause >nul