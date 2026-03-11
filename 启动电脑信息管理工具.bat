@echo off
cd /d "%~dp0"
for %%F in ("%~dp0*.ps1") do (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%%~fF"
    goto :eof
)
echo PowerShell script not found.
pause
