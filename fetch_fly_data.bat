@echo off
setlocal

rem Fetch Fly.io volume data to local machine.
rem Requirements: flyctl installed and logged in (`fly auth login`).

set "APP=akare-emisyon"
set "REMOTE_ARCHIVE=/tmp/data-backup.tar.gz"
set "OUT_DIR=%~dp0data-downloaded"
set "OUT_TAR=%~dp0data-backup.tar.gz"

echo [1/5] Creating archive on Fly volume...
fly ssh console -a %APP% -C "tar czf %REMOTE_ARCHIVE% /data"
if errorlevel 1 goto :err

echo [2/5] Downloading archive to %OUT_TAR% ...
fly sftp get %REMOTE_ARCHIVE% "%OUT_TAR%"
if errorlevel 1 goto :err

echo [3/5] Cleaning remote temp archive...
fly ssh console -a %APP% -C "rm -f %REMOTE_ARCHIVE%"

echo [4/5] Preparing local extract directory %OUT_DIR% ...
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"

echo [5/5] Extracting archive...
tar xzf "%OUT_TAR%" -C "%OUT_DIR%"
if errorlevel 1 goto :err

echo Done. Local data is in:
echo   %OUT_DIR%\data
echo To run the app locally with this data in the current shell:
echo   set DATA_DIR=%OUT_DIR%\data
echo   python app.py
goto :eof

:err
echo FAILED. Check flyctl login and network, then retry.
exit /b 1

endlocal
