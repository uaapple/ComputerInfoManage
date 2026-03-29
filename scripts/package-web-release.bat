@echo off
cd /d "%~dp0\.."
powershell -ExecutionPolicy Bypass -File ".\scripts\package-web-release.ps1"
if errorlevel 1 (
  echo.
  echo Package failed.
  pause
  exit /b 1
)
echo.
echo Package completed.
pause
