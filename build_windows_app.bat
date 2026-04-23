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

echo [4/4] Creating distribution folder structure...
if not exist "dist\QBOX\Data"           mkdir "dist\QBOX\Data"
if not exist "dist\QBOX\Data\Backups"   mkdir "dist\QBOX\Data\Backups"
if not exist "dist\QBOX\ATTACHMENTS"    mkdir "dist\QBOX\ATTACHMENTS"

echo [5/4] Build complete.
echo.
echo Distribution folder:  dist\QBOX\
echo.
echo   dist\QBOX\QBOX.exe          - run the app
echo   dist\QBOX\Data\             - put query_tracker.xlsx here
echo   dist\QBOX\Data\Backups\     - auto-backups saved here
echo   dist\QBOX\ATTACHMENTS\      - attachment files per query
echo.
echo Share the entire dist\QBOX folder with your team.
exit /b 0
