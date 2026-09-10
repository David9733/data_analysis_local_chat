@echo off
cd /d "%~dp0"

if exist "%USERPROFILE%\.local\bin\python.exe" (
    start "" "%USERPROFILE%\.local\bin\python.exe" local_chat.py
    exit /b 0
)

if exist "%USERPROFILE%\.local\bin\python3.14.exe" (
    start "" "%USERPROFILE%\.local\bin\python3.14.exe" local_chat.py
    exit /b 0
)

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    start "" python local_chat.py
    exit /b 0
)

echo Python is not found. Please install Python.
pause
