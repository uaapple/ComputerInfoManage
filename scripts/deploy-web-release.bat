@echo off
cd /d "%~dp0\.."
setlocal

set PACKAGE_PATH=%~1
set DEFAULT_PACKAGE_DIR=C:\Apps\Deploy\ComputerInfoManage
set ELEVATED_FLAG=%~2

if /I "%~1"=="ELEVATED" (
  set PACKAGE_PATH=
  set ELEVATED_FLAG=ELEVATED
)

net session >nul 2>&1
if errorlevel 1 (
  if /I not "%ELEVATED_FLAG%"=="ELEVATED" (
    echo.
    echo Requesting administrator permission...
    if "%PACKAGE_PATH%"=="" (
      powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList @('ELEVATED') -Verb RunAs"
    ) else (
      powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList @('%PACKAGE_PATH%','ELEVATED') -Verb RunAs"
    )
    if errorlevel 1 (
      echo.
      echo Administrator permission was not granted.
      pause
      exit /b 1
    )
    exit /b 0
  )
  echo.
  echo Administrator permission is required to deploy the service.
  pause
  exit /b 1
)

if "%PACKAGE_PATH%"=="" (
  echo.
  echo Deploy Web Release
  echo.
  echo Usage:
  echo   1. Drag a release zip onto this bat file
  echo   2. Or paste / drag the zip path into this window, then press Enter
  echo   3. Or just press Enter to use the newest zip from:
  echo      %DEFAULT_PACKAGE_DIR%
  echo.
  set /p PACKAGE_PATH=Package zip path: 
)

if "%PACKAGE_PATH%"=="" (
  echo.
  for /f "delims=" %%I in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "$dir = '%DEFAULT_PACKAGE_DIR%'; if (Test-Path -LiteralPath $dir) { Get-ChildItem -LiteralPath $dir -Filter *.zip | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName }"') do set PACKAGE_PATH=%%I
)

if "%PACKAGE_PATH%"=="" (
  echo.
  echo No package path was provided, and no zip package was found in:
  echo   %DEFAULT_PACKAGE_DIR%
  pause
  exit /b 1
)

if not exist "%PACKAGE_PATH%" (
  echo.
  echo Package not found:
  echo   %PACKAGE_PATH%
  pause
  exit /b 1
)

echo.
echo Using package:
echo   %PACKAGE_PATH%
echo.

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
