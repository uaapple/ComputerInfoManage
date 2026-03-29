[CmdletBinding()]
param(
  [string]$InstallRoot = "D:\ITK-ComputerInfoManage",
  [string]$ServiceName = "ITKComputerInfoManageWeb",
  [string]$NodePath = "",
  [string]$NssmPath = "",
  [string]$HostAddress = "0.0.0.0",
  [int]$Port = 8099
)

$ErrorActionPreference = "Stop"

function Resolve-NodePath {
  param([string]$PreferredPath)

  if ($PreferredPath -and (Test-Path -LiteralPath $PreferredPath)) {
    return (Resolve-Path $PreferredPath).Path
  }

  $command = Get-Command node -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }

  $commonPath = "C:\Program Files\nodejs\node.exe"
  if (Test-Path -LiteralPath $commonPath) {
    return $commonPath
  }

  throw "Node.js not found. Please install Node.js or pass -NodePath."
}

function Resolve-NssmPath {
  param([string]$PreferredPath)

  if ($PreferredPath -and (Test-Path -LiteralPath $PreferredPath)) {
    return (Resolve-Path $PreferredPath).Path
  }

  $command = Get-Command nssm -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }

  $candidate = Join-Path $InstallRoot "tools\nssm\nssm.exe"
  if (Test-Path -LiteralPath $candidate) {
    return $candidate
  }

  throw "nssm.exe not found. Please install NSSM and pass -NssmPath, or place it at $candidate."
}

$nodeExe = Resolve-NodePath -PreferredPath $NodePath
$nssmExe = Resolve-NssmPath -PreferredPath $NssmPath
$appRoot = Join-Path $InstallRoot "app\current"
$appScript = Join-Path $appRoot "web-server.js"
$dataRoot = Join-Path $InstallRoot "data"
$logsRoot = Join-Path $InstallRoot "logs"
$serviceExists = [bool](Get-Service -Name $ServiceName -ErrorAction SilentlyContinue)

if (-not (Test-Path -LiteralPath $appScript)) {
  throw "Missing deployed app at $appScript. Please deploy the release first."
}

New-Item -ItemType Directory -Path $dataRoot -Force | Out-Null
New-Item -ItemType Directory -Path $logsRoot -Force | Out-Null

if (-not $serviceExists) {
  & $nssmExe install $ServiceName $nodeExe "web-server.js"
}

& $nssmExe set $ServiceName AppDirectory $appRoot
& $nssmExe set $ServiceName DisplayName $ServiceName
& $nssmExe set $ServiceName Description "ITK China ComputerInfoManage Web"
& $nssmExe set $ServiceName Start SERVICE_AUTO_START
& $nssmExe set $ServiceName AppStdout (Join-Path $logsRoot "service-stdout.log")
& $nssmExe set $ServiceName AppStderr (Join-Path $logsRoot "service-stderr.log")
& $nssmExe set $ServiceName AppRotateFiles 1
& $nssmExe set $ServiceName AppRotateOnline 1
& $nssmExe set $ServiceName AppEnvironmentExtra "HOST=$HostAddress" "PORT=$Port" "DATA_DIR=$dataRoot"

Start-Service -Name $ServiceName

Write-Host "Service installed: $ServiceName"
Write-Host "Node path: $nodeExe"
Write-Host "NSSM path: $nssmExe"
