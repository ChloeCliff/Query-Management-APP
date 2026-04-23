@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  QBOX - First-time setup
echo ==========================================
echo.
echo  Run this once after extracting QBOX from SharePoint.
echo.

if not exist "Data"             mkdir "Data"
if not exist "Data\Backups"     mkdir "Data\Backups"
if not exist "ATTACHMENTS"      mkdir "ATTACHMENTS"

echo Folders created:
echo   Data\            ^ place your query_tracker.xlsx here
echo   Data\Backups\    ^ daily backups will be saved here automatically
echo   ATTACHMENTS\     ^ attachment files per query
echo.

if exist "Data\query_tracker.xlsx" (
  echo query_tracker.xlsx already found in Data\ - ready to launch.
) else (
  echo NEXT STEP: Copy your query_tracker.xlsx into the Data\ folder,
  echo            then launch QBOX.exe.
)

echo.
pause
