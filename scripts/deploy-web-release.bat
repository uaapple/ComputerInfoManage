@echo off
cd /d "%~dp0\.."

set PACKAGE_PATH=%~1

if "%PACKAGE_PATH%"=="" (
  echo.
  echo Usage:
  echo   deploy-web-release.bat "D:\path\to\release.zip"
  echo.
  echo Or drag a release zip onto this bat file.
  pause
  exit /b 1
)

powershell -ExecutionPolicy Bypass -File ".\scripts\deploy-web-release.ps1" -PackagePath "%PACKAGE_PATH%"
if errorlevel 1 (
  echo.
  echo Deploy failed.
  pause
  exit /b 1
)
echo.
echo Deploy completed.
pause
