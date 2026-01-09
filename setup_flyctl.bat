@echo off
rem One-time installer for flyctl on Windows

echo Installing flyctl...
powershell -Command "iwr https://fly.io/install.ps1 -useb | iex"

echo Updating PATH for flyctl (may require new terminal)...
setx PATH "%PATH%;%USERPROFILE%\.fly\bin"

echo.
echo Done. Now open a NEW PowerShell window and run:
echo   flyctl version
echo   fly auth login
echo Then you can double-click sync_volume_and_deploy.bat to sync and deploy.
pause
