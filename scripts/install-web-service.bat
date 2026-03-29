@echo off
cd /d "%~dp0\.."

set INSTALL_ROOT=D:\ITK-ComputerInfoManage
set NSSM_PATH=D:\tools\nssm\nssm.exe

echo.
echo This will install or update the Windows service.
echo Current InstallRoot: %INSTALL_ROOT%
echo Current NSSM path:   %NSSM_PATH%
echo If needed, edit scripts\install-web-service.bat before double-clicking.
echo.

powershell -ExecutionPolicy Bypass -File ".\scripts\install-web-service.ps1" -InstallRoot "%INSTALL_ROOT%" -NssmPath "%NSSM_PATH%"
if errorlevel 1 (
  echo.
  echo Service install failed.
  pause
  exit /b 1
)
echo.
echo Service install completed.
pause
