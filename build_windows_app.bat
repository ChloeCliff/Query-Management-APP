@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Building QBOX Windows application...
echo ==========================================

set "PY_CMD="
set "ICON_ARGS="
set "DATA_ARGS="
set "RELEASE_DIR=QBOX"
set "BUILD_ARCHIVE_DIR=%RELEASE_DIR%\_build"

where py >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3"
if "%PY_CMD%"=="" (
  where python >nul 2>nul
  if not errorlevel 1 set "PY_CMD=python"
)
if "%PY_CMD%"=="" (
  echo [ERROR] Python not found.
  echo Install Python 3.11+ and ensure it is on PATH.
  exit /b 1
)

echo [1/6] Installing build dependencies...
%PY_CMD% -m pip install --disable-pip-version-check --no-input --upgrade pip
if errorlevel 1 (
  echo [ERROR] Failed to upgrade pip.
  exit /b 1
)
%PY_CMD% -m pip install --disable-pip-version-check --no-input pyinstaller openpyxl pyspellchecker pillow cairosvg
if errorlevel 1 (
  echo [ERROR] Failed to install dependencies.
  exit /b 1
)

echo [2/6] Cleaning previous build output...
if exist build      rmdir /s /q build
if exist dist       rmdir /s /q dist
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
if exist QBOX.spec  del /q QBOX.spec

echo [3/6] Generating icon assets...
if exist "qbox-logo.svg" (
  echo     - Running generate_icons.py
  %PY_CMD% generate_icons.py
  if errorlevel 1 (
    echo [ERROR] Failed to generate icon assets.
    exit /b 1
  )
  if exist "qbox.ico" set "ICON_ARGS=--icon qbox.ico"
  set "DATA_ARGS=--add-data qbox-logo.svg;."
) else (
  echo [WARNING] qbox-logo.svg not found. Building without custom icon assets.
)
if exist "qbox-icon-64.png"  set "DATA_ARGS=%DATA_ARGS% --add-data qbox-icon-64.png;."
if exist "qbox-icon-128.png" set "DATA_ARGS=%DATA_ARGS% --add-data qbox-icon-128.png;."
if exist "qbox-icon-256.png" set "DATA_ARGS=%DATA_ARGS% --add-data qbox-icon-256.png;."

echo [4/6] Running PyInstaller one-file build...
%PY_CMD% -m PyInstaller --noconfirm --clean --onefile --windowed --name QBOX --collect-data spellchecker --hidden-import spellchecker.resources %DATA_ARGS% %ICON_ARGS% app.py
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo [5/6] Creating release folder structure...
mkdir "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%\ATTACHMENTS"
mkdir "%RELEASE_DIR%\Data"
mkdir "%RELEASE_DIR%\Data\Backups"
mkdir "%BUILD_ARCHIVE_DIR%"

copy /y "dist\QBOX.exe" "%RELEASE_DIR%\QBOX.exe" >nul
if errorlevel 1 (
  echo [ERROR] Could not copy QBOX.exe into the release folder.
  exit /b 1
)

if exist "setup_folders.bat" copy /y "setup_folders.bat" "%RELEASE_DIR%\setup_folders.bat" >nul
if exist "QBOX.spec" copy /y "QBOX.spec" "%BUILD_ARCHIVE_DIR%\QBOX.spec" >nul
if exist "dist\QBOX.exe" copy /y "dist\QBOX.exe" "%BUILD_ARCHIVE_DIR%\QBOX.exe" >nul

echo [6/6] Build final checks...
if not exist "%RELEASE_DIR%\QBOX.exe" (
  echo [ERROR] Release exe was not generated.
  exit /b 1
)

echo.
echo ==========================================
echo  Build complete
echo ==========================================
echo.
echo  Release folder: %RELEASE_DIR%\
echo.
echo    %RELEASE_DIR%\QBOX.exe
echo    %RELEASE_DIR%\ATTACHMENTS\
echo    %RELEASE_DIR%\Data\
echo    %RELEASE_DIR%\Data\Backups\
echo    %RELEASE_DIR%\_build\
echo.
echo  Share the entire QBOX folder with each team.
echo  The app will keep query_tracker.xlsx and backups in Data\.
echo ==========================================
exit /b 0
