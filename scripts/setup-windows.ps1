$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectDir

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js 22.13 or newer is required. Install it from your approved software source, then run this script again."
}

node -e 'const [major, minor] = process.versions.node.split(".").map(Number); if (major < 22 || (major === 22 && minor < 13)) { console.error(`Node ${process.versions.node} is too old; Presentation Studio requires Node 22.13 or newer.`); process.exit(1); }'
if ($LASTEXITCODE -ne 0) { throw "The installed Node.js version is unsupported." }

Write-Host "Installing locked source dependencies..."
npm ci
npm test
npm run check:data-safety
npm run build

Write-Host ""
Write-Host "Presentation Studio source setup passed."
Write-Host "Start it with: $PSScriptRoot\start-windows.ps1"
Write-Host "MCP configuration: node $PSScriptRoot\configure-mcp.mjs"
