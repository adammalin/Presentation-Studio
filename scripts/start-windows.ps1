$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectDir

if (-not (Test-Path (Join-Path $ProjectDir "node_modules"))) {
  throw "Presentation Studio is not set up. Run scripts\setup-windows.ps1 first."
}

npm start
