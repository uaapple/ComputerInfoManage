[CmdletBinding()]
param(
  [string]$PackagePath = "",
  [string]$PackageRoot = "",
  [string]$InstallRoot = "D:\ITK-ComputerInfoManage",
  [string]$ServiceName = "ITKComputerInfoManageWeb",
  [switch]$SkipServiceRestart
)

$ErrorActionPreference = "Stop"

function Ensure-Directory {
  param([Parameter(Mandatory = $true)][string]$Path)

  New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Stop-ServiceIfExists {
  param([Parameter(Mandatory = $true)][string]$Name)

  $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
  if (-not $service) {
    return $false
  }

  if ($service.Status -ne "Stopped") {
    Stop-Service -Name $Name -Force
    $service.WaitForStatus("Stopped", [TimeSpan]::FromSeconds(30))
  }

  return $true
}

function Start-ServiceIfExists {
  param([Parameter(Mandatory = $true)][string]$Name)

  $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
  if (-not $service) {
    return $false
  }

  Start-Service -Name $Name
  $service = Get-Service -Name $Name
  $service.WaitForStatus("Running", [TimeSpan]::FromSeconds(30))
  return $true
}

function Initialize-DataFiles {
  param([Parameter(Mandatory = $true)][string]$DataRoot)

  $jsonFiles = @(
    "computers.json",
    "colleagues.json",
    "inventory_mail_batches.json"
  )

  foreach ($jsonFile in $jsonFiles) {
    $target = Join-Path $DataRoot $jsonFile
    if (-not (Test-Path -LiteralPath $target)) {
      Set-Content -LiteralPath $target -Value "[]`n" -Encoding UTF8
    }
  }
}

if (-not $PackageRoot -and -not $PackagePath) {
  $PackageRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$tempExtractRoot = $null

if ($PackagePath) {
  $resolvedPackagePath = (Resolve-Path $PackagePath).Path
  if ($resolvedPackagePath.EndsWith(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
    $tempExtractRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ComputerInfoManageDeploy-" + [guid]::NewGuid().ToString("N"))
    Expand-Archive -LiteralPath $resolvedPackagePath -DestinationPath $tempExtractRoot -Force
    $PackageRoot = $tempExtractRoot
  } else {
    $PackageRoot = $resolvedPackagePath
  }
}

$PackageRoot = (Resolve-Path $PackageRoot).Path
$packageAppRoot = Join-Path $PackageRoot "app"

if (-not (Test-Path -LiteralPath (Join-Path $packageAppRoot "web-server.js"))) {
  throw "Package root is invalid. Missing app\web-server.js."
}

$appBaseRoot = Join-Path $InstallRoot "app"
$currentRoot = Join-Path $appBaseRoot "current"
$incomingRoot = Join-Path $appBaseRoot "incoming"
$backupRoot = Join-Path $InstallRoot "backups"
$dataRoot = Join-Path $InstallRoot "data"
$logsRoot = Join-Path $InstallRoot "logs"
$packagesRoot = Join-Path $InstallRoot "packages"
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupTarget = Join-Path $backupRoot $timestamp

Ensure-Directory -Path $InstallRoot
Ensure-Directory -Path $appBaseRoot
Ensure-Directory -Path $backupRoot
Ensure-Directory -Path $dataRoot
Ensure-Directory -Path $logsRoot
Ensure-Directory -Path $packagesRoot
Initialize-DataFiles -DataRoot $dataRoot

if ($PackagePath) {
  Copy-Item -LiteralPath $resolvedPackagePath -Destination (Join-Path $packagesRoot ([System.IO.Path]::GetFileName($resolvedPackagePath))) -Force
}

$serviceExists = Stop-ServiceIfExists -Name $ServiceName

if (Test-Path -LiteralPath $incomingRoot) {
  Remove-Item -LiteralPath $incomingRoot -Recurse -Force
}

Copy-Item -LiteralPath $packageAppRoot -Destination $incomingRoot -Recurse

if (Test-Path -LiteralPath $currentRoot) {
  Ensure-Directory -Path $backupTarget
  Move-Item -LiteralPath $currentRoot -Destination (Join-Path $backupTarget "current")
}

Move-Item -LiteralPath $incomingRoot -Destination $currentRoot

if (-not $SkipServiceRestart -and $serviceExists) {
  Start-ServiceIfExists -Name $ServiceName | Out-Null
}

Write-Host "Deploy completed."
Write-Host "InstallRoot: $InstallRoot"
Write-Host "Current app: $currentRoot"
Write-Host "Data root:   $dataRoot"
Write-Host "Logs root:   $logsRoot"
if ($serviceExists) {
  Write-Host "Service '$ServiceName' restarted."
} else {
  Write-Host "Service '$ServiceName' not found. Run scripts\install-web-service.ps1 on the VM if needed."
}

if ($tempExtractRoot -and (Test-Path -LiteralPath $tempExtractRoot)) {
  Remove-Item -LiteralPath $tempExtractRoot -Recurse -Force
}
