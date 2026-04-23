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

echo [1/5] Installing build dependencies...
py -3 -m pip install --upgrade pip >nul
py -3 -m pip install pyinstaller openpyxl pyspellchecker
if errorlevel 1 (
  echo [ERROR] Failed to install dependencies.
  exit /b 1
)

echo [2/5] Cleaning previous build output...
if exist build      rmdir /s /q build
if exist dist       rmdir /s /q dist
if exist QBOX.spec  del /q QBOX.spec
if exist QBOX.zip   del /q QBOX.zip

echo [3/5] Running PyInstaller...
py -3 -m PyInstaller --noconfirm --clean --windowed --name QBOX app.py
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo [4/5] Creating distribution folder structure...
if not exist "dist\QBOX\Data"           mkdir "dist\QBOX\Data"
if not exist "dist\QBOX\Data\Backups"   mkdir "dist\QBOX\Data\Backups"
if not exist "dist\QBOX\ATTACHMENTS"    mkdir "dist\QBOX\ATTACHMENTS"

rem Copy the first-run setup helper into the distribution
copy /y "setup_folders.bat" "dist\QBOX\setup_folders.bat" >nul

echo [5/5] Creating QBOX.zip for SharePoint upload...
powershell -NoProfile -Command "Compress-Archive -Path 'dist\QBOX\*' -DestinationPath 'QBOX.zip' -Force"
if errorlevel 1 (
  echo [WARNING] Could not create QBOX.zip (PowerShell error). Upload dist\QBOX\ manually.
) else (
  echo.
  echo  QBOX.zip created successfully.
)

echo.
echo ==========================================
echo  Build complete
echo ==========================================
echo.
echo  Upload QBOX.zip to your team SharePoint.
echo  Team members:
echo    1. Download QBOX.zip from SharePoint
echo    2. Extract to a local folder (e.g. C:\QBOX)
echo    3. Run setup_folders.bat once (creates Data\ and ATTACHMENTS\)
echo    4. Copy your query_tracker.xlsx into Data\
echo    5. Launch QBOX.exe
echo.
echo  Daily backups are saved automatically to Data\Backups\.
echo ==========================================
exit /b 0
