@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo  QBOX - First-time folder setup
echo ==========================================
echo.

if not exist "Data"             mkdir "Data"
if not exist "Data\Backups"     mkdir "Data\Backups"
if not exist "ATTACHMENTS"      mkdir "ATTACHMENTS"

echo Folders ready:
echo   Data\            - place your query_tracker.xlsx here
echo   Data\Backups\    - daily backups will be saved here automatically
echo   ATTACHMENTS\     - drop attachment files here per query
echo.

if exist "Data\query_tracker.xlsx" (
  echo query_tracker.xlsx already exists - no changes made.
) else (
  echo NOTE: Copy your query_tracker.xlsx into the Data\ folder before
  echo       launching QBOX for the first time.
)

echo.
echo Setup complete. Launch QBOX.exe to start.
pause
