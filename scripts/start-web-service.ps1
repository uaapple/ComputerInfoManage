[CmdletBinding()]
param(
  [string]$ServiceName = "ITKComputerInfoManageWeb"
)

$ErrorActionPreference = "Stop"

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
  throw "Service '$ServiceName' was not found."
}

if ($service.Status -eq "Running") {
  Write-Host "Service '$ServiceName' is already running."
  exit 0
}

Start-Service -Name $ServiceName
Write-Host "Service '$ServiceName' started."
