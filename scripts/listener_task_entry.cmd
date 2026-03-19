@echo off
setlocal
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
title OpenClaw WeChat Listener
cd /d "%~dp0.."
set "LOG_FILE=%USERPROFILE%\.openclaw\logs\wechat-listener-task.log"
set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"
echo [%date% %time%] starting listener task entry with "%PYTHON_EXE%"
"%PYTHON_EXE%" -u scripts\listener_task_entry.py %*
set "EXIT_CODE=%errorlevel%"
echo [%date% %time%] listener task entry exited with code %EXIT_CODE%
>> "%LOG_FILE%" echo [%date% %time%] listener task entry exited with code %EXIT_CODE%
echo.
echo Listener exited. Press any key to close this window.
pause >nul
