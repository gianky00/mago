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
:: Use 'where' to find all python.exe in PATH and pick the first valid one
for /f "delims=" %%i in ('where python 2^>nul') do (
    if not defined PYTHON_EXE (
        echo "%%i" | find /I "\WindowsApps\" >nul
        if errorlevel 1 (
            set "PYTHON_EXE=%%~i"
        )
    )
)

:: Fallback if not found in PATH or if 'where' only found invalid ones
if not defined PYTHON_EXE (
    echo Fallback: Searching in common installation directories...
    for /r "%ProgramFiles%" %%f in (python.exe) do (
        if not defined PYTHON_EXE (
            echo "%%f" | find /I "\WindowsApps\" >nul
            if errorlevel 1 ( set "PYTHON_EXE=%%~f" )
        )
    )
    for /r "%ProgramFiles(x86)%" %%f in (python.exe) do (
        if not defined PYTHON_EXE (
            echo "%%f" | find /I "\WindowsApps\" >nul
            if errorlevel 1 ( set "PYTHON_EXE=%%~f" )
        )
    )
    for /r "%LocalAppData%\Programs\Python" %%f in (python.exe) do (
        if not defined PYTHON_EXE (
            echo "%%f" | find /I "\WindowsApps\" >nul
            if errorlevel 1 ( set "PYTHON_EXE=%%~f" )
        )
    )
)

:found_python
if not defined PYTHON_EXE (
    echo.
    echo ERROR: Could not find a valid Python installation.
    echo Make sure a real version of Python is installed and in the PATH.
    goto :error
)

echo Found Python at: %PYTHON_EXE%
echo.

:: ============================================================================
::  Create and activate the virtual environment
:: ============================================================================
if not exist "%~dp0%VENV_NAME%\Scripts\activate.bat" (
    echo Creating virtual environment '%VENV_NAME%'...
    
    :: Pre-create the directory to avoid race conditions with AV/system
    if not exist "%~dp0%VENV_NAME%" mkdir "%~dp0%VENV_NAME%"
    
    "%PYTHON_EXE%" -m venv "%~dp0%VENV_NAME%" >nul 2>&1

    :: Verify creation by checking for the activate script
    if not exist "%~dp0%VENV_NAME%\Scripts\activate.bat" (
        echo ERROR: Failed to create virtual environment. The activation script was not found.
        goto :error
    )
    echo Virtual environment created successfully.
) else (
    echo Virtual environment '%VENV_NAME%' already exists.
)

echo Activating the virtual environment...
call "%~dp0%VENV_NAME%\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment.
    goto :error
)
echo.

:: ============================================================================
::  Upgrade core packaging tools
:: ============================================================================
echo Upgrading pip, setuptools, and wheel...
python.exe -m pip install --upgrade pip setuptools wheel --no-cache-dir
if %errorlevel% neq 0 (
    echo ERROR: Failed to upgrade packaging tools.
    goto :error
)
echo Core tools upgraded successfully.
echo.

:: ============================================================================
::  Install dependencies
:: ============================================================================
echo Installing dependencies from '%REQUIREMENTS_FILE%'...
echo (Using --no-cache-dir to avoid permission issues)
python.exe -m pip install -r "%REQUIREMENTS_FILE%" --no-cache-dir
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    goto :error
)
echo Dependencies installed successfully.
echo.

:: ============================================================================
::  Check for and Install Tesseract OCR
:: ============================================================================
echo.
echo Checking for Tesseract OCR installation...

set "TESSERACT_DIR=%ProgramFiles%\Tesseract-OCR"
set "TESSERACT_EXE=%TESSERACT_DIR%\tesseract.exe"
set "TESSERACT_URL=https://github.com/UB-Mannheim/tesseract/releases/download/v5.2.0/tesseract-ocr-w64-setup-v5.2.0.20220712.exe"
set "TESSERACT_INSTALLER=%TEMP%\tesseract-installer.exe"

if exist "%TESSERACT_EXE%" (
    echo Tesseract is already installed at '%TESSERACT_EXE%'.
) else (
    echo Tesseract not found. Attempting to download and install automatically...

    echo Downloading Tesseract installer using Python...
    python.exe download_util.py "%TESSERACT_URL%" "%TESSERACT_INSTALLER%"
    if %errorlevel% neq 0 (
        echo ERROR: The Python download script failed.
        goto :error
    )

    if not exist "%TESSERACT_INSTALLER%" (
        echo ERROR: Download script ran but the installer file was not created.
        echo Please install it manually from: %TESSERACT_URL%
        goto :error
    )

    :: Verify the downloaded file size is reasonable (e.g., > 1MB)
    for %%F in ("%TESSERACT_INSTALLER%") do set "FILESIZE=%%~zF"
    if %FILESIZE% LSS 1048576 (
        echo ERROR: Downloaded file is too small (%FILESIZE% bytes^). It is likely corrupt or incomplete.
        echo Deleting invalid file: %TESSERACT_INSTALLER%
        del "%TESSERACT_INSTALLER%"
        echo Please check your network/firewall settings and try again.
        goto :error
    )
    echo Downloaded file size is %FILESIZE% bytes. Looks good.

    echo Running Tesseract installer. This may take a moment...
    :: Use 'start' to launch the installer in a separate, non-blocking process.
    :: This avoids both the deadlock from 'start /wait' and the race condition from direct execution.
    start "" "%TESSERACT_INSTALLER%" /S /AllUsers /D="%ProgramFiles%\Tesseract-OCR"
    if %errorlevel% neq 0 (
        echo ERROR: Tesseract installer failed to start.
        goto :error
    )

    echo Waiting for installation to complete (will check for up to 90 seconds)...
    set "TIMEOUT_SECONDS=90"
    set "COUNT=0"
    :wait_for_tesseract_loop
    if exist "%TESSERACT_EXE%" (
        echo Tesseract executable found.
        goto :tesseract_install_finished
    )

    if %COUNT% GEQ %TIMEOUT_SECONDS% (
        echo ERROR: Tesseract installation timed out after %TIMEOUT_SECONDS% seconds.
        echo The installer may have failed or is still running. Please check the Task Manager.
        del "%TESSERACT_INSTALLER%"
        goto :error
    )

    :: Wait for 2 seconds before checking again
    timeout /t 2 /nobreak >nul
    set /a COUNT+=2
    goto :wait_for_tesseract_loop

    :tesseract_install_finished
    :: Brief pause to ensure file handles are released before proceeding
    timeout /t 2 /nobreak >nul

    del "%TESSERACT_INSTALLER%"

    if not exist "%TESSERACT_EXE%" (
        echo ERROR: Tesseract installation failed. The executable was not found at the default location.
        goto :error
    )
    echo Tesseract installed successfully.
)

echo.
echo Updating Tesseract path in config.json...
python.exe update_config.py "%TESSERACT_EXE%"
if %errorlevel% neq 0 (
    echo WARNING: Failed to update Tesseract path in config.json.
) else (
    echo Configuration updated successfully.
)
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
pause

:end
echo.
echo Press any key to close the window.
pause