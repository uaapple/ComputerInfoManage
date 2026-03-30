[CmdletBinding()]
param(
  [string]$OutputDir = "",
  [string]$Version = "",
  [switch]$IncludeSeedData
)

$ErrorActionPreference = "Stop"

function New-CleanDirectory {
  param([Parameter(Mandatory = $true)][string]$Path)

  if (Test-Path -LiteralPath $Path) {
    Remove-Item -LiteralPath $Path -Recurse -Force
  }

  New-Item -ItemType Directory -Path $Path | Out-Null
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$OutputDir = if ($OutputDir) { $OutputDir } else { Join-Path $repoRoot "release" }
$packageJsonPath = Join-Path $repoRoot "package.json"
$packageJson = Get-Content -LiteralPath $packageJsonPath -Raw | ConvertFrom-Json

$baseVersion = if ($Version) { $Version } else { [string]$packageJson.version }
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$releaseName = "ComputerInfoManage-web-$baseVersion-$timestamp"
$releaseRoot = Join-Path $OutputDir $releaseName
$appRoot = Join-Path $releaseRoot "app"
$docsRoot = Join-Path $releaseRoot "docs"
$scriptsRoot = Join-Path $releaseRoot "scripts"
$seedDataRoot = Join-Path $releaseRoot "seed-data"
$zipPath = Join-Path $OutputDir "$releaseName.zip"

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-CleanDirectory -Path $releaseRoot
New-Item -ItemType Directory -Path $appRoot | Out-Null
New-Item -ItemType Directory -Path $docsRoot | Out-Null
New-Item -ItemType Directory -Path $scriptsRoot | Out-Null

Copy-Item -LiteralPath (Join-Path $repoRoot "web-server.js") -Destination $appRoot
Copy-Item -LiteralPath (Join-Path $repoRoot "package.json") -Destination $appRoot
Copy-Item -LiteralPath (Join-Path $repoRoot "ITK_Logo_RGB.jpg") -Destination $appRoot
Copy-Item -LiteralPath (Join-Path $repoRoot "web") -Destination $appRoot -Recurse

Copy-Item -LiteralPath (Join-Path $repoRoot "README.md") -Destination $releaseRoot
Copy-Item -LiteralPath (Join-Path $repoRoot "PROJECT_CONTEXT.md") -Destination $releaseRoot
Copy-Item -LiteralPath (Join-Path $repoRoot "docs\requirements.md") -Destination $docsRoot
Copy-Item -LiteralPath (Join-Path $repoRoot "docs\deployment.md") -Destination $docsRoot

$scriptFiles = @(
  "deploy-web-release.ps1",
  "deploy-web-release.bat",
  "install-web-service.ps1",
  "install-web-service.bat",
  "start-web-service.ps1",
  "start-web-service.bat",
  "stop-web-service.ps1",
  "stop-web-service.bat",
  "package-web-release.bat"
)

foreach ($scriptFile in $scriptFiles) {
  Copy-Item -LiteralPath (Join-Path $PSScriptRoot $scriptFile) -Destination $scriptsRoot
}

if ($IncludeSeedData) {
  New-Item -ItemType Directory -Path $seedDataRoot | Out-Null
  Copy-Item -LiteralPath (Join-Path $repoRoot "data\computers.json") -Destination $seedDataRoot
  Copy-Item -LiteralPath (Join-Path $repoRoot "data\colleagues.json") -Destination $seedDataRoot
  Copy-Item -LiteralPath (Join-Path $repoRoot "data\inventory_mail_batches.json") -Destination $seedDataRoot
}

$versionText = @(
  "ReleaseName=$releaseName"
  "Version=$baseVersion"
  "BuildTime=$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))"
)

Set-Content -LiteralPath (Join-Path $releaseRoot "VERSION.txt") -Value $versionText -Encoding UTF8

$manifest = [ordered]@{
  releaseName = $releaseName
  version = $baseVersion
  buildTime = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  includeSeedData = [bool]$IncludeSeedData
  appFiles = @(
    "app\web-server.js",
    "app\package.json",
    "app\ITK_Logo_RGB.jpg",
    "app\web\index.html",
    "app\web\styles.css",
    "app\web\app.js"
  )
}

$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $releaseRoot "release-manifest.json") -Encoding UTF8

if (Test-Path -LiteralPath $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $releaseRoot "*") -DestinationPath $zipPath -Force

Write-Host "Release folder created: $releaseRoot"
Write-Host "Release zip created:    $zipPath"
