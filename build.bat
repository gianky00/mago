@echo off
setlocal

:: ============================================================================
::  Configuration
:: ============================================================================
set "BUILD_DIR=build"
set "DIST_DIR=dist"
set "MAIN_SCRIPT=gui.py"
set "LICENSE_FILE=license.lic"

:: ============================================================================
::  Build Process
:: ============================================================================
echo Starting the build process...

:: 1. Clean up previous build artifacts
if exist "%BUILD_DIR%" (
    echo Deleting old build directory...
    rmdir /s /q "%BUILD_DIR%"
)
if exist "%DIST_DIR%" (
    echo Deleting old dist directory...
    rmdir /s /q "%DIST_DIR%"
)

:: 2. Create the build directory
echo Creating build directory...
mkdir "%BUILD_DIR%"

:: 3. Copy necessary files to the build directory
echo Copying application files...
copy "*.py" "%BUILD_DIR%\"
copy "*.json" "%BUILD_DIR%\"
copy "*.db" "%BUILD_DIR%\"
copy "*.xlsm" "%BUILD_DIR%\"
xcopy "file di setup" "%BUILD_DIR%\file di setup\" /E /I /Q

:: 4. Obfuscate the code using PyArmor
echo Obfuscating scripts...
pyarmor obfuscate --recursive --output "%DIST_DIR%" "%BUILD_DIR%\%MAIN_SCRIPT%"

:: 5. Copy the license file
if exist "%LICENSE_FILE%" (
    echo Copying license file...
    copy "%LICENSE_FILE%" "%DIST_DIR%\"
) else (
    echo WARNING: license.lic not found. The obfuscated application will not run without it.
)

:: 6. Create the launcher for the obfuscated application
echo Creating launcher script...
(
    echo @echo off
    echo setlocal
    echo REM This script runs the obfuscated application. It assumes 'python.exe' is in the system's PATH.
    echo python.exe "%~dp0%MAIN_SCRIPT%"
    echo endlocal
    echo pause
) > "%DIST_DIR%\avvio_obfuscated.bat"

echo.
echo Build process finished successfully!
echo The obfuscated application is located in the '%DIST_DIR%' directory.
echo Use 'avvio_obfuscated.bat' to run it.
echo.

endlocal
pause
