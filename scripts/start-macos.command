#!/bin/zsh
set -euo pipefail

SCRIPT_DIR=${0:A:h}
PROJECT_DIR=${SCRIPT_DIR:h}
cd "$PROJECT_DIR"

if [[ ! -d node_modules ]]; then
  print -u2 "Presentation Studio is not set up. Run scripts/setup-macos.command first."
  exit 1
fi

npm start
