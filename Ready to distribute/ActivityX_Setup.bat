@echo off
setlocal enabledelayedexpansion

echo ========================================
echo ActivityX Automatic Setup
echo ========================================
echo.

REM Get current directory
set "SETUP_DIR=%~dp0"
cd /d "%SETUP_DIR%"

REM Define installation directory (LocalAppData for reliability)
set "INSTALL_DIR=%LOCALAPPDATA%\ActivityX"

echo Step 1: Running initial setup...
echo ----------------------------------------
echo Running Python and dependency installation (this may take a moment)...

REM Run setup.bat content inline to avoid the pause issue
REM Get current directory
set "SETUP_INSTALL_DIR=%~dp0"
cd /d "%SETUP_INSTALL_DIR%"

REM Check if Python is installed
echo    Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo    Python not found. Installing Python...
    if exist "python-installer.exe" (
        python-installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
        echo    Python installed successfully
    ) else (
        echo    WARNING: python-installer.exe not found
    )
) else (
    echo    Python already installed
)

REM Install Python dependencies
echo    Installing Python dependencies...
if exist "requirements.txt" (
    pip install -r requirements.txt --quiet
    echo    Dependencies installed
) else (
    echo    WARNING: requirements.txt not found, skipping dependencies
)

echo    Initial setup completed!

echo.
echo Step 2: Creating installation directory...
echo ----------------------------------------
echo    Target directory: %INSTALL_DIR%
if not exist "%INSTALL_DIR%" (
    echo    Creating directory...
    mkdir "%INSTALL_DIR%"
    if exist "%INSTALL_DIR%" (
        echo    Successfully created: %INSTALL_DIR%
    ) else (
        echo    ERROR: Failed to create directory
    )
) else (
    echo    Directory already exists: %INSTALL_DIR%
)

echo.
echo Step 3: Copying files to installation directory...
echo ----------------------------------------
if exist "activity_tracker.exe" (
    copy "activity_tracker.exe" "%INSTALL_DIR%\" >nul
    if exist "%INSTALL_DIR%\activity_tracker.exe" (
        echo    Successfully copied: activity_tracker.exe
    ) else (
        echo    ERROR: Failed to copy activity_tracker.exe
    )
) else (
    echo    ERROR: activity_tracker.exe not found in current directory
)

if exist "activity_tracker_controller.exe" (
    copy "activity_tracker_controller.exe" "%INSTALL_DIR%\" >nul
    if exist "%INSTALL_DIR%\activity_tracker_controller.exe" (
        echo    Successfully copied: activity_tracker_controller.exe
    ) else (
        echo    ERROR: Failed to copy activity_tracker_controller.exe
    )
) else (
    echo    ERROR: activity_tracker_controller.exe not found in current directory
)

if exist "keytrk_data" (
    xcopy "keytrk_data" "%INSTALL_DIR%\keytrk_data\" /E /I /Y >nul
    if exist "%INSTALL_DIR%\keytrk_data" (
        echo    Successfully copied: keytrk_data folder
    ) else (
        echo    ERROR: Failed to copy keytrk_data folder
    )
) else (
    echo    WARNING: keytrk_data not found in current directory, skipping
)

echo.
echo Step 4: Creating startup shortcuts...
echo ----------------------------------------

REM Get startup folder path
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

REM Create VBS script to create shortcuts
set "VBS_SCRIPT=%TEMP%\create_shortcuts.vbs"
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%VBS_SCRIPT%"
echo sLinkFile1 = "%STARTUP_FOLDER%\ActivityTracker.lnk" >> "%VBS_SCRIPT%"
echo Set oLink1 = oWS.CreateShortcut(sLinkFile1) >> "%VBS_SCRIPT%"
echo oLink1.TargetPath = "%INSTALL_DIR%\activity_tracker.exe" >> "%VBS_SCRIPT%"
echo oLink1.WorkingDirectory = "%INSTALL_DIR%" >> "%VBS_SCRIPT%"
echo oLink1.Description = "ActivityX Tracker" >> "%VBS_SCRIPT%"
echo oLink1.Save >> "%VBS_SCRIPT%"
echo sLinkFile2 = "%STARTUP_FOLDER%\ActivityTrackerController.lnk" >> "%VBS_SCRIPT%"
echo Set oLink2 = oWS.CreateShortcut(sLinkFile2) >> "%VBS_SCRIPT%"
echo oLink2.TargetPath = "%INSTALL_DIR%\activity_tracker_controller.exe" >> "%VBS_SCRIPT%"
echo oLink2.WorkingDirectory = "%INSTALL_DIR%" >> "%VBS_SCRIPT%"
echo oLink2.Description = "ActivityX Tracker Controller" >> "%VBS_SCRIPT%"
echo oLink2.Save >> "%VBS_SCRIPT%"

REM Execute VBS script
cscript //nologo "%VBS_SCRIPT%"
if %errorlevel% equ 0 (
    echo    Created startup shortcuts successfully
) else (
    echo    WARNING: Failed to create startup shortcuts
)

REM Clean up temporary VBS script
del "%VBS_SCRIPT%" >nul 2>&1

echo.
echo Step 5: Starting applications...
echo ----------------------------------------
cd /d "%INSTALL_DIR%"
start "" "activity_tracker.exe"
echo    Started: activity_tracker.exe
timeout /t 2 /nobreak >nul
start "" "activity_tracker_controller.exe"
echo    Started: activity_tracker_controller.exe
cd /d "%SETUP_DIR%"

echo.
echo Step 6: Registering Windows Scheduled Task...
echo ----------------------------------------
REM Delete existing task if present
schtasks /Delete /TN "ActivityX Controller" /F >nul 2>&1
REM Create scheduled task: runs every 5 minutes + on logon
schtasks /Create /TN "ActivityX Controller" /TR "\"%INSTALL_DIR%\activity_tracker_controller.exe\"" /SC MINUTE /MO 5 /F >nul 2>&1
if %errorlevel% equ 0 (
    echo    Scheduled task created: runs every 5 minutes
) else (
    echo    WARNING: Failed to create scheduled task (may need admin rights)
)
REM Also create a logon trigger
schtasks /Create /TN "ActivityX Controller Startup" /TR "\"%INSTALL_DIR%\activity_tracker_controller.exe\"" /SC ONLOGON /F >nul 2>&1
if %errorlevel% equ 0 (
    echo    Logon task created: starts on user login
) else (
    echo    WARNING: Failed to create logon task
)

echo.
echo ========================================
echo INSTALLATION COMPLETE!
echo ========================================
echo.
echo Installation location: %INSTALL_DIR%
echo.
echo The following applications have been installed:
echo - ActivityX Tracker
echo - ActivityX Tracker Controller
echo.
echo Both applications will now start automatically when Windows starts.
echo The applications are currently running in the background.
echo.
echo You can find the installed files at:
echo %INSTALL_DIR%
echo.
pause