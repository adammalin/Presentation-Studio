#!/bin/zsh
set -euo pipefail

SCRIPT_DIR=${0:A:h}
PROJECT_DIR=${SCRIPT_DIR:h}
cd "$PROJECT_DIR"

if ! command -v node >/dev/null 2>&1; then
  print -u2 "Node.js 22.13 or newer is required. Install it from your approved software source, then run this script again."
  exit 1
fi

node -e 'const [major, minor] = process.versions.node.split(".").map(Number); if (major < 22 || (major === 22 && minor < 13)) { console.error(`Node ${process.versions.node} is too old; Presentation Studio requires Node 22.13 or newer.`); process.exit(1); }'

print "Installing locked source dependencies..."
npm ci
print "Running tests and data-safety checks..."
npm test
npm run check:data-safety
npm run build

print ""
print "Presentation Studio source setup passed."
print "Start it with: $PROJECT_DIR/scripts/start-macos.command"
print "MCP configuration: node $PROJECT_DIR/scripts/configure-mcp.mjs"
