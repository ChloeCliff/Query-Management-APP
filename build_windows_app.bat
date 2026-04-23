@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Building QBOX Windows application...
echo ==========================================

where py >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python launcher (py) not found.
  echo Install Python 3.11+ from python.org and retry.
  exit /b 1
)

echo [1/4] Installing build dependencies...
py -3 -m pip install --upgrade pip >nul
py -3 -m pip install pyinstaller openpyxl pyspellchecker
if errorlevel 1 (
  echo [ERROR] Failed to install dependencies.
  exit /b 1
)

echo [2/4] Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist QBOX.spec del /q QBOX.spec

echo [3/4] Running PyInstaller...
py -3 -m PyInstaller --noconfirm --clean --windowed --name QBOX app.py
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo [4/4] Build complete.
echo.
echo App folder:
echo   dist\QBOX\
echo.
echo Exe to run:
echo   dist\QBOX\QBOX.exe
echo.
echo Share the entire dist\QBOX folder with your teammate.
exit /b 0
