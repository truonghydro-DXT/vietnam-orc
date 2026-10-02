@echo off
title GEOAI Server Launcher
set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%"

echo.
echo ================================
echo   Kich hoat moi truong Python
echo ================================
echo.

call "%ROOT_DIR%.venv\Scripts\activate.bat"

if errorlevel 1 (
    echo ❌ Loi: Khong the kich hoat moi truong venv.
    pause
    exit /b
)

echo.
echo ================================
echo   Chay Flask server
echo ================================
echo.

"%ROOT_DIR%.venv\Scripts\python.exe" ui.py

if errorlevel 1 (
    echo ❌ Loi khi chay ui.py
    pause
    exit /b
)

echo.
echo ----------------
echo ✅ Server da ket thuc.
echo ----------------
pause