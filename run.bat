@echo off
setlocal enabledelayedexpansion

:: ========================================
::   wechat-skill - OpenClaw WeChat Integration
:: ========================================

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

:: Parse command line arguments
set "ACTION="
set "DURATION="
set "NO_NARRATOR="

:parse_args
if "%~1"=="" goto end_parse
if /i "%~1"=="start" set "ACTION=start"
if /i "%~1"=="stop" set "ACTION=stop"
if /i "%~1"=="env" set "ACTION=env"
if /i "%~1"=="listen" set "ACTION=listen"
if /i "%~1"=="send" set "ACTION=send"
if /i "%~1"=="relay" set "ACTION=relay"
if /i "%~1"=="status" set "ACTION=status"
if /i "%~1"=="install" set "ACTION=install"
if /i "%~1"=="--duration" set "DURATION=%~2" & shift
if /i "%~1"=="--no-narrator" set "NO_NARRATOR=1"
shift
goto parse_args
:end_parse

:: Show menu if no action specified
if "%ACTION%"=="" goto show_menu

:: Execute action
if "%ACTION%"=="start" goto do_env
if "%ACTION%"=="stop" goto do_stop
if "%ACTION%"=="env" goto do_env
if "%ACTION%"=="listen" goto do_listen
if "%ACTION%"=="send" goto do_send
if "%ACTION%"=="relay" goto do_relay
if "%ACTION%"=="status" goto do_status
if "%ACTION%"=="install" goto do_install
goto show_menu

:: ========================================
::   Interactive Menu
:: ========================================
:show_menu
cls
echo.
echo  ====================================================
echo       wechat-skill - OpenClaw WeChat Integration
echo  ====================================================
echo.
echo   1. Prepare Environment (Start Narrator + WeChat)
echo   2. Start Listening Service
echo   3. Stop Listening Service
echo   4. Send Message
echo   5. Pending Relay Management
echo   6. Check Status
echo   7. Install Dependencies
echo   8. Exit
echo.
set /p choice=Select option (1-8):

if "%choice%"=="1" goto do_env
if "%choice%"=="2" goto do_listen
if "%choice%"=="3" goto do_stop
if "%choice%"=="4" goto do_send
if "%choice%"=="5" goto do_relay
if "%choice%"=="6" goto do_status
if "%choice%"=="7" goto do_install
if "%choice%"=="8" goto end
echo [ERROR] Invalid option
timeout /t 2 >nul
goto show_menu

:: ========================================
::   Environment Preparation
:: ========================================
:do_env
cls
echo.
echo  ====================================================
echo   Environment Preparation
echo  ====================================================
echo.

:: Check Python
echo [1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.9+
    pause
    goto show_menu
)
echo       Python OK

:: Check pyweixin
echo [2/4] Checking pyweixin...
python -c "from pyweixin import Navigator" >nul 2>&1
if errorlevel 1 (
    echo [WARN] pyweixin not installed. Run option 7 to install.
    echo       Or: pip install -e /path/to/pywechat
) else (
    echo       pyweixin OK
)

:: Check WeChat
echo [3/4] Checking WeChat...
tasklist /fi "imagename eq WeChat.exe" 2>nul | find /i "WeChat.exe" >nul
if errorlevel 1 (
    echo       WeChat not running, starting...
    start "" "WeChat.exe" 2>nul
    if errorlevel 1 (
        echo [WARN] Cannot start WeChat automatically. Please start manually.
    ) else (
        echo       Waiting for WeChat to start...
        timeout /t 5 >nul
    )
) else (
    echo       WeChat is running
)

:: Start Narrator
if "%NO_NARRATOR%"=="1" (
    echo [4/4] Skipping Narrator (--no-narrator flag)
) else (
    echo [4/4] Starting Narrator...
    start "" "C:\Windows\System32\narrator.exe" 2>nul
    echo       Narrator started (will auto-close in 5 minutes)

    :: Create delayed stop script
    echo @echo off > "%TEMP%\stop_narrator_delayed.bat"
    echo timeout /t 300 ^>nul >> "%TEMP%\stop_narrator_delayed.bat"
    echo taskkill /f /im narrator.exe ^>nul 2^>^&1 >> "%TEMP%\stop_narrator_delayed.bat"
    echo exit >> "%TEMP%\stop_narrator_delayed.bat"
    start "" /min cmd /c "%TEMP%\stop_narrator_delayed.bat"
)

echo.
echo  ====================================================
echo   Environment Ready!
echo  ====================================================
echo.
pause
goto show_menu

:: ========================================
::   Start Listening
:: ========================================
:do_listen
cls
echo.
echo  ====================================================
echo   Start Listening Service
echo  ====================================================
echo.

:: Check config
if not exist "config.yaml" (
    echo [WARN] config.yaml not found, creating from example...
    copy "config.example.yaml" "config.yaml" >nul 2>&1
    echo       Created config.yaml, please edit it before starting.
    echo.
    notepad "config.yaml"
    pause
    goto show_menu
)

:: Get duration
if "%DURATION%"=="" set /p DURATION=Enter duration (default 24h):
if "%DURATION%"=="" set "DURATION=24h"

echo.
echo Starting listener with duration: %DURATION%
echo Press Ctrl+C to stop
echo.

python scripts\wechat_listener.py --config config.yaml --duration %DURATION% --no-close
pause
goto show_menu

:: ========================================
::   Stop Listening
:: ========================================
:do_stop
cls
echo.
echo  ====================================================
echo   Stop Listening Service
echo  ====================================================
echo.

echo Stopping wechat_listener processes...
taskkill /f /im python.exe /fi "windowtitle eq *wechat_listener*" 2>nul
wmic process where "commandline like '%%wechat_listener%%'" delete 2>nul

echo Stopping Narrator...
taskkill /f /im narrator.exe >nul 2>&1

echo.
echo [DONE] All services stopped.
pause
goto show_menu

:: ========================================
::   Send Message
:: ========================================
:do_send
cls
echo.
echo  ====================================================
echo   Send Message
echo  ====================================================
echo.
set /p send_group=Enter group/friend name:
set /p send_msg=Enter message:
set /p send_at=Enter members to @ (comma-separated, optional):

if "%send_at%"=="" (
    python scripts\send_wechat_message.py --group "%send_group%" --message "%send_msg%"
) else (
    python scripts\send_wechat_message.py --group "%send_group%" --message "%send_msg%" --at "%send_at%"
)
echo.
pause
goto show_menu

:: ========================================
::   Pending Relay Management
:: ========================================
:do_relay
cls
echo.
echo  ====================================================
echo   Pending Relay Management
echo  ====================================================
echo.
echo   1. Record pending relay
echo   2. Query pending relay
echo   3. Clear pending relay
echo   4. Back to main menu
echo.
set /p relay_choice=Select option (1-4):

if "%relay_choice%"=="1" goto relay_record
if "%relay_choice%"=="2" goto relay_query
if "%relay_choice%"=="3" goto relay_clear
if "%relay_choice%"=="4" goto show_menu
goto do_relay

:relay_record
echo.
set /p relay_target=Enter target group name:
set /p relay_source=Enter source group name:
set /p relay_at=Enter original sender:
python scripts\record_pending_relay.py --supplier-group "%relay_target%" --to-group "%relay_source%" --at "%relay_at%"
echo.
pause
goto do_relay

:relay_query
echo.
set /p relay_query_group=Enter target group name:
python scripts\get_pending_relay.py --supplier-group "%relay_query_group%"
if errorlevel 1 echo [INFO] No pending relay found.
echo.
pause
goto do_relay

:relay_clear
echo.
set /p relay_clear_group=Enter target group name:
python scripts\get_pending_relay.py --supplier-group "%relay_clear_group%" --clear
echo [DONE] Pending relay cleared.
echo.
pause
goto do_relay

:: ========================================
::   Check Status
:: ========================================
:do_status
cls
echo.
echo  ====================================================
echo   System Status
echo  ====================================================
echo.

:: Python
echo [Python]
python --version 2>&1
echo.

:: pyweixin
echo [pyweixin]
python -c "from pyweixin import Navigator; print('  Installed OK')" 2>&1
echo.

:: WeChat
echo [WeChat]
tasklist /fi "imagename eq WeChat.exe" 2>nul | find /i "WeChat.exe" >nul
if errorlevel 1 (
    echo   Not running
) else (
    echo   Running
)
echo.

:: Narrator
echo [Narrator]
tasklist /fi "imagename eq narrator.exe" 2>nul | find /i "narrator.exe" >nul
if errorlevel 1 (
    echo   Not running
) else (
    echo   Running
)
echo.

:: Listener
echo [Listener]
wmic process where "commandline like '%%wechat_listener%%'" get processid 2>nul | find /i "python" >nul
if errorlevel 1 (
    echo   Not running
) else (
    echo   Running
)
echo.

:: Config
echo [Config]
if exist "config.yaml" (
    echo   config.yaml: OK
) else (
    echo   config.yaml: Not found
)
if exist "wechat_pending_relay.json" (
    echo   pending relay: Has data
) else (
    echo   pending relay: Empty
)
echo.

pause
goto show_menu

:: ========================================
::   Install Dependencies
:: ========================================
:do_install
cls
echo.
echo  ====================================================
echo   Install Dependencies
echo  ====================================================
echo.

echo [1/3] Installing requirements.txt...
pip install -r requirements.txt
echo.

echo [2/3] Checking pyweixin...
python -c "from pyweixin import Navigator" >nul 2>&1
if errorlevel 1 (
    echo pyweixin not installed.
    echo.
    echo Please install pyweixin manually:
    echo   pip install -e /path/to/pywechat
    echo.
    echo Or copy pywechat folder to this directory:
    echo   %SCRIPT_DIR%lib\pywechat\
    echo.
    set /p pywechat_path=Enter pywechat path (or press Enter to skip):

    if not "!pywechat_path!"=="" (
        echo       Installing from !pywechat_path!...
        pip install -e "!pywechat_path!"
        python -c "from pyweixin import Navigator" >nul 2>&1
        if errorlevel 1 (
            echo [ERROR] Installation failed. Please check the path.
        ) else (
            echo       pyweixin installed successfully!
        )
    )
) else (
    echo pyweixin already installed.
)
echo.

echo [3/3] Done!
echo.
pause
goto show_menu

:: ========================================
::   Exit
:: ========================================
:end
cls
echo.
echo  Thank you for using wechat-skill!
echo.
echo  Remember: Do not use for illegal activities!
echo.
pause
exit /b 0
