@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

echo.
echo  ========================================
echo   wechat-skill Setup Script
echo  ========================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

:: Check Python
echo [Step 1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Please install Python 3.9+ from https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VER=%%i
echo       Python %PYTHON_VER% found.

:: Install pip packages
echo.
echo [Step 2/4] Installing Python packages...
pip install pyyaml pyautogui pywinauto emoji --quiet
if errorlevel 1 (
    echo [WARN] Some packages may not have installed correctly.
) else (
    echo       Core packages installed.
)

:: Check pyweixin
echo.
echo [Step 3/4] Checking pyweixin...
python -c "from pyweixin import Navigator" >nul 2>&1
if errorlevel 1 (
    echo       pyweixin not installed.
    echo.
    echo  ----------------------------------------
    echo   pyweixin Installation Required
    echo  ----------------------------------------
    echo.
    echo   pyweixin is required for this skill to work.
    echo   You have two options:
    echo.
    echo   Option 1: Install from local path
    echo     pip install -e D:\WorkSpace\python\pywechat
    echo.
    echo   Option 2: Install from the pywechat folder in this directory
    echo     If you have pywechat source code, copy it to:
    echo     %SCRIPT_DIR%pywechat\
    echo.
    echo   Then run: pip install -e %SCRIPT_DIR%pywechat
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
    echo       pyweixin already installed.
)

:: Create config file
echo.
echo [Step 4/4] Creating config file...
if not exist "config.yaml" (
    if exist "config.example.yaml" (
        copy "config.example.yaml" "config.yaml" >nul
        echo       Created config.yaml from example.
        echo.
        echo  ========================================
        echo   IMPORTANT: Please edit config.yaml
        echo  ========================================
        echo.
        echo   1. Set your OpenClaw webhook URL and token
        echo   2. Add the WeChat groups you want to monitor
        echo.
        notepad "config.yaml"
    ) else (
        echo [WARN] config.example.yaml not found.
    )
) else (
    echo       config.yaml already exists.
)

:: Summary
echo.
echo  ========================================
echo   Setup Complete!
echo  ========================================
echo.
echo   Next steps:
echo   1. Make sure WeChat is logged in
echo   2. Run: run.bat
echo   3. Select option 1 to prepare environment
echo   4. Select option 2 to start listening
echo.
echo   Documentation:
echo     - README.md  : Project overview
echo     - USAGE.md   : Detailed usage guide
echo.
pause
