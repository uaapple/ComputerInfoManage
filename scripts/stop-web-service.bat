@echo off
cd /d "%~dp0\.."
powershell -ExecutionPolicy Bypass -File ".\scripts\stop-web-service.ps1"
if errorlevel 1 (
  echo.
  echo Stop service failed.
  pause
  exit /b 1
)
echo.
echo Service stopped.
pause
