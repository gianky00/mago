@echo off
setlocal

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
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
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
:: Search for python.exe in the PATH, excluding the Windows Store path
for %%i in (python.exe) do (
    set "FOUND_PATH=%%~$PATH:i"
    if defined FOUND_PATH (
        echo "%%FOUND_PATH%%" | find /I "\WindowsApps\" >nul
        if errorlevel 1 (
            set "PYTHON_EXE=%%~$PATH:i"
            goto found_python
        )
    )
)

:: Fallback if not found in PATH (e.g., installations not added to PATH)
if not defined PYTHON_EXE (
    for /r "%ProgramFiles%" %%f in (python.exe) do (
        echo "%%f" | find /I "\WindowsApps\" >nul
        if errorlevel 1 (
            set "PYTHON_EXE=%%f"
            goto found_python
        )
    )
    for /r "%LocalAppData%\Programs\Python" %%f in (python.exe) do (
         echo "%%f" | find /I "\WindowsApps\" >nul
        if errorlevel 1 (
            set "PYTHON_EXE=%%f"
            goto found_python
        )
    )
)

:found_python
if not defined PYTHON_EXE (
    echo.
    echo ERROR: Could not find a valid Python installation.
    echo Make sure Python is installed and, if possible, added to the PATH.
    goto :error
)

echo Found Python at: %PYTHON_EXE%
echo.

:: ============================================================================
::  Create and activate the virtual environment
:: ============================================================================
if not exist "%VENV_NAME%\Scripts\activate.bat" (
    echo Creating virtual environment '%VENV_NAME%'...
    "%PYTHON_EXE%" -m venv %VENV_NAME%
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment.
        goto :error
    )
    echo Virtual environment created.
) else (
    echo Virtual environment '%VENV_NAME%' already exists.
)

echo Activating the virtual environment...
call "%VENV_NAME%\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment.
    goto :error
)
echo.

:: ============================================================================
::  Install dependencies
:: ============================================================================
echo Installing dependencies from '%REQUIREMENTS_FILE%'...
echo (Using --no-cache-dir to avoid permission issues)
pip install -r "%REQUIREMENTS_FILE%" --no-cache-dir
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    goto :error
)
echo Dependencies installed successfully.
echo.

:: ============================================================================
::  Check for Tesseract installation
:: ============================================================================
set "TESSERACT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe"
if not exist "%TESSERACT_PATH%" (
    echo.
    echo WARNING: Tesseract OCR does not appear to be installed at '%TESSERACT_PATH%'.
    echo The automation may not work correctly if OCR popups are required.
    echo Please install Tesseract from https://github.com/UB-Mannheim/tesseract/wiki
    echo and ensure it is installed in the default path.
    echo.
)

:: ============================================================================
::  Run the main Python script
:: ============================================================================
echo Starting the script '%PYTHON_SCRIPT%'...
echo.
python "%PYTHON_SCRIPT%"
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
pause

:end
echo.
echo Press any key to close the window.
pause