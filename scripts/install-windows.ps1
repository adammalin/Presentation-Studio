$ErrorActionPreference = "Stop"

$Repository = "adammalin/Presentation-Studio"
$Branch = "codex/web-slide-design-engine"
$NodeVersion = "v22.13.0"
$InstallRoot = if ($env:PRESENTATION_STUDIO_INSTALL_DIR) {
  $env:PRESENTATION_STUDIO_INSTALL_DIR
} else {
  Join-Path $env:LOCALAPPDATA "Presentation Studio"
}
$AppDir = Join-Path $InstallRoot "app"
$RuntimeRoot = Join-Path $InstallRoot "runtime"
$RuntimeNodeDir = Join-Path $RuntimeRoot "node"
$StagingDir = Join-Path $InstallRoot (".app-staging-" + $PID)
$PreviousDir = Join-Path $InstallRoot "app.previous"
$TempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("presentation-studio-install-" + [Guid]::NewGuid().ToString("N"))

function Stop-Install([string]$Message) {
  throw "Presentation Studio installation failed: $Message"
}

function Test-CompatibleNode {
  $Node = Get-Command node.exe -ErrorAction SilentlyContinue
  $Npm = Get-Command npm.cmd -ErrorAction SilentlyContinue
  if (-not $Node -or -not $Npm) { return $false }
  & $Node.Source -e 'const [major, minor] = process.versions.node.split(".").map(Number); process.exit(major > 22 || (major === 22 && minor >= 13) ? 0 : 1)' *> $null
  return $LASTEXITCODE -eq 0
}

function Invoke-Checked([string]$Executable, [string[]]$Arguments) {
  & $Executable @Arguments
  if ($LASTEXITCODE -ne 0) {
    Stop-Install "$Executable exited with code $LASTEXITCODE."
  }
}

if ([string]::IsNullOrWhiteSpace($InstallRoot) -or
    $InstallRoot -eq [System.IO.Path]::GetPathRoot($InstallRoot) -or
    $InstallRoot -eq $env:USERPROFILE) {
  Stop-Install "the install location is not safe. Set PRESENTATION_STUDIO_INSTALL_DIR to a dedicated folder."
}

if ((Test-Path $AppDir) -and -not (Test-Path (Join-Path $AppDir ".presentation-studio-managed-install"))) {
  Stop-Install "$AppDir already exists and is not managed by this installer. Move it or choose another PRESENTATION_STUDIO_INSTALL_DIR."
}

New-Item -ItemType Directory -Force -Path $InstallRoot, $RuntimeRoot, $TempDir | Out-Null

try {
  if (Test-CompatibleNode) {
    Write-Host "Using Node.js $(& node.exe --version) from the current system."
  } else {
    $Architecture = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString().ToLowerInvariant()
    $NodeArchitecture = switch ($Architecture) {
      "x64" { "x64" }
      "arm64" { "arm64" }
      default { Stop-Install "unsupported Windows architecture $Architecture." }
    }
    $ArchiveName = "node-$NodeVersion-win-$NodeArchitecture.zip"
    $ArchivePath = Join-Path $TempDir $ArchiveName
    $ChecksumsPath = Join-Path $TempDir "SHASUMS256.txt"

    Write-Host "Downloading the portable Node.js prerequisite..."
    Invoke-WebRequest -UseBasicParsing -Uri "https://nodejs.org/dist/$NodeVersion/$ArchiveName" -OutFile $ArchivePath
    Invoke-WebRequest -UseBasicParsing -Uri "https://nodejs.org/dist/$NodeVersion/SHASUMS256.txt" -OutFile $ChecksumsPath
    $ChecksumLine = Get-Content $ChecksumsPath | Where-Object { $_ -match "\s+$([Regex]::Escape($ArchiveName))$" } | Select-Object -First 1
    if (-not $ChecksumLine) { Stop-Install "Node.js did not publish a checksum for $ArchiveName." }
    $ExpectedChecksum = ($ChecksumLine -split "\s+")[0].ToLowerInvariant()
    $ActualChecksum = (Get-FileHash -Algorithm SHA256 $ArchivePath).Hash.ToLowerInvariant()
    if ($ActualChecksum -ne $ExpectedChecksum) { Stop-Install "the Node.js archive checksum did not match the official manifest." }

    $NodeExtractRoot = Join-Path $TempDir "node"
    Expand-Archive -Path $ArchivePath -DestinationPath $NodeExtractRoot -Force
    $ExtractedNodeDir = Join-Path $NodeExtractRoot "node-$NodeVersion-win-$NodeArchitecture"
    if (-not (Test-Path (Join-Path $ExtractedNodeDir "node.exe"))) { Stop-Install "the downloaded Node.js runtime is incomplete." }
    if (Test-Path $RuntimeNodeDir) { Remove-Item -Recurse -Force -LiteralPath $RuntimeNodeDir }
    Move-Item -LiteralPath $ExtractedNodeDir -Destination $RuntimeNodeDir
    $env:PATH = "$RuntimeNodeDir;$env:PATH"
  }

  if (-not (Test-CompatibleNode)) { Stop-Install "Node.js $NodeVersion or newer could not be prepared." }
  $NpmCommand = (Get-Command npm.cmd -ErrorAction Stop).Source

  $SourceArchiveUrl = if ($env:PRESENTATION_STUDIO_SOURCE_ARCHIVE_URL) {
    $env:PRESENTATION_STUDIO_SOURCE_ARCHIVE_URL
  } else {
    "https://github.com/$Repository/archive/refs/heads/$Branch.zip"
  }
  $SourceArchivePath = Join-Path $TempDir "presentation-studio.zip"
  $SourceExtractRoot = Join-Path $TempDir "source"
  Write-Host "Downloading the latest Presentation Studio 0.2.1 source..."
  Invoke-WebRequest -UseBasicParsing -Uri $SourceArchiveUrl -OutFile $SourceArchivePath
  Expand-Archive -Path $SourceArchivePath -DestinationPath $SourceExtractRoot -Force
  $SourceDir = Get-ChildItem -LiteralPath $SourceExtractRoot -Directory | Select-Object -First 1
  if (-not $SourceDir -or -not (Test-Path (Join-Path $SourceDir.FullName "package-lock.json"))) {
    Stop-Install "the downloaded source archive is not a Presentation Studio release."
  }
  Move-Item -LiteralPath $SourceDir.FullName -Destination $StagingDir

  Write-Host "Installing locked dependencies and verifying the staged application..."
  Push-Location $StagingDir
  try {
    Invoke-Checked $NpmCommand @("ci")
    Invoke-Checked $NpmCommand @("test")
    Invoke-Checked $NpmCommand @("run", "check:data-safety")
    Invoke-Checked $NpmCommand @("run", "build")
    Set-Content -LiteralPath (Join-Path $StagingDir ".presentation-studio-managed-install") -Value "managed by scripts/install-windows.ps1" -Encoding ASCII
  } finally {
    Pop-Location
  }

  if (Test-Path $PreviousDir) {
    if (-not (Test-Path (Join-Path $PreviousDir ".presentation-studio-managed-install"))) {
      Stop-Install "$PreviousDir is not a managed backup and will not be replaced."
    }
    Remove-Item -Recurse -Force -LiteralPath $PreviousDir
  }
  if (Test-Path $AppDir) { Move-Item -LiteralPath $AppDir -Destination $PreviousDir }
  Move-Item -LiteralPath $StagingDir -Destination $AppDir

  $Launcher = Join-Path $InstallRoot "Launch Presentation Studio.cmd"
  $LauncherContent = @'
@echo off
setlocal
set "INSTALL_ROOT=%~dp0"
if exist "%INSTALL_ROOT%runtime\node\node.exe" set "PATH=%INSTALL_ROOT%runtime\node;%PATH%"
cd /d "%INSTALL_ROOT%app"
call npm start
'@
  Set-Content -LiteralPath $Launcher -Value $LauncherContent -Encoding ASCII

  Write-Host ""
  Write-Host "Presentation Studio 0.2.1 installed successfully."
  Write-Host "Install location: $AppDir"
  Write-Host "Launcher: $Launcher"
  Write-Host "MCP configuration command: node `"$AppDir\scripts\configure-mcp.mjs`""

  if ($env:PRESENTATION_STUDIO_NO_LAUNCH -ne "1") {
    Write-Host "Starting Presentation Studio..."
    Start-Process -FilePath $Launcher
  }
} finally {
  if (Test-Path $StagingDir) { Remove-Item -Recurse -Force -LiteralPath $StagingDir }
  if (Test-Path $TempDir) { Remove-Item -Recurse -Force -LiteralPath $TempDir }
}
