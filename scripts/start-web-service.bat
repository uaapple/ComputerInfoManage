@echo off
cd /d "%~dp0\.."
powershell -ExecutionPolicy Bypass -File ".\scripts\start-web-service.ps1"
if errorlevel 1 (
  echo.
  echo Start service failed.
  pause
  exit /b 1
)
echo.
echo Service started.
pause
