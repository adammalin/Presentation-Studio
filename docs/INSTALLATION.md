# Presentation Studio source installation

Presentation Studio is distributed as source during the initial release. There is no DMG, PKG, MSI, EXE, app-store package, or other application installer.

## Requirements

- macOS or Windows
- Node.js 22.13 or newer
- npm
- Local disk space for source decks, self-contained project packages, and exported copies

## macOS

1. Download or clone the repository to a local folder.
2. Double-click `scripts/setup-macos.command`, or run it from Terminal.
3. The script runs `npm ci`, tests, the repository data-safety scan, and a production build.
4. Start Presentation Studio with `scripts/start-macos.command`.

The script does not change Gatekeeper or other operating-system protections.

## Windows

1. Download or clone the repository to a local folder.
2. Open PowerShell in the repository.
3. Run `& .\scripts\setup-windows.ps1`.
4. Start Presentation Studio with `& .\scripts\start-windows.ps1`.

The script does not change PowerShell execution policy, SmartScreen, or other operating-system protections. If local policy prevents scripts from running, use your organization's approved process rather than disabling protections.

## MCP clients

Run `node scripts/configure-mcp.mjs` to print the standard `mcpServers` entry. Add the `presentation-studio` entry to any MCP-capable client's configuration. To write a selected JSON configuration explicitly, use:

```sh
node scripts/configure-mcp.mjs --write /absolute/path/to/mcp-config.json
```

Presentation Studio must be open. The MCP server uses STDIO and connects to the active app through a per-session token on a loopback-only bridge. The in-app AI session switch is off by default. MCP can stage a proposal but cannot apply, save, or export it.

## Project files and encryption

- `.pstudio` is a ZIP-based self-contained package with canonical project JSON and immutable-by-hash Resources.
- `.pstudio-secure` encrypts the complete package with AES-256-GCM and PBKDF2-SHA-256.
- Presentation Studio cannot recover a lost project password.
- Project encryption does not cover the external original files or separately exported PowerPoint, PDF, SVG, or PNG files.

## Verify the source checkout

Run:

```sh
npm test
npm run check:data-safety
npm run build
npm run desktop:smoke
```

A passing build proves that the current source compiles and the Electron renderer opens. It does not approve a cleaned deck for distribution; exported decks remain subject to human production, brand, content, and technical review.
