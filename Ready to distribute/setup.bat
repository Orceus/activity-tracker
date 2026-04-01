@echo off
echo ========================================
echo KeyTRK Activity Tracker Setup
echo ========================================
echo.

REM Get current directory
set "INSTALL_DIR=%~dp0"
cd /d "%INSTALL_DIR%"

REM Check if Python is installed
echo 1. Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo    Python not found. Installing Python...
    if exist "python-installer.exe" (
        python-installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
        echo    Python installed successfully
    ) else (
        echo    ERROR: python-installer.exe not found
        pause
        exit /b 1
    )
) else (
    echo    Python already installed
)

REM Install Python dependencies
echo 2. Installing Python dependencies...
if exist "requirements.txt" (
    pip install -r requirements.txt --quiet
    echo    Dependencies installed
) else (
    echo    WARNING: requirements.txt not found, skipping dependencies
)

echo.
echo ========================================
echo SETUP COMPLETE!
echo ========================================
echo.
echo.
pause