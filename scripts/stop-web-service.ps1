[CmdletBinding()]
param(
  [string]$ServiceName = "ITKComputerInfoManageWeb"
)

$ErrorActionPreference = "Stop"

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
  throw "Service '$ServiceName' was not found."
}

if ($service.Status -eq "Stopped") {
  Write-Host "Service '$ServiceName' is already stopped."
  exit 0
}

Stop-Service -Name $ServiceName -Force
Write-Host "Service '$ServiceName' stopped."
