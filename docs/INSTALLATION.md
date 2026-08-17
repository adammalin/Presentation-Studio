# Presentation Studio source installation

Presentation Studio is distributed as source during the initial release. There is no DMG, PKG, MSI, EXE, app-store package, or other application installer.

This guide installs Presentation Studio 0.2.0 from the isolated `codex/web-slide-design-engine` branch. It does not merge or change the repository's `main` branch.

## Requirements

- macOS or Windows
- Node.js 22.13 or newer
- npm
- Local disk space for source decks, self-contained project packages, and exported copies

## Get the current source release

For a fresh checkout:

```sh
git clone --branch codex/web-slide-design-engine --single-branch https://github.com/adammalin/Presentation-Studio.git
cd Presentation-Studio
```

To update an existing Git checkout, close Presentation Studio and run:

```sh
git fetch origin codex/web-slide-design-engine
git switch codex/web-slide-design-engine
git pull --ff-only origin codex/web-slide-design-engine
```

The fast-forward-only pull stops instead of overwriting local development changes. After cloning or updating, run the platform setup script below so dependencies and the production renderer match the checked-out version.

## macOS

1. Get or update the release branch using the commands above.
2. Double-click `scripts/setup-macos.command`, or run it from Terminal in the repository.
3. The script runs `npm ci`, tests, the repository data-safety scan, and a production build.
4. Start Presentation Studio with `scripts/start-macos.command`.

The script does not change Gatekeeper or other operating-system protections.

## Windows

1. Get or update the release branch using the commands above.
2. Open PowerShell in the repository.
3. Run `& .\scripts\setup-windows.ps1`.
4. Start Presentation Studio with `& .\scripts\start-windows.ps1`.

The script does not change PowerShell execution policy, SmartScreen, or other operating-system protections. If local policy prevents scripts from running, use your organization's approved process rather than disabling protections.

## MCP clients

Run `node scripts/configure-mcp.mjs` to print the standard `mcpServers` entry. Add the `presentation-studio` entry to any MCP-capable client's configuration. To write a selected JSON configuration explicitly, use:

```sh
node scripts/configure-mcp.mjs --write /absolute/path/to/mcp-config.json
```

Presentation Studio must be open. The MCP server uses STDIO and connects to the active app through a per-session token on a loopback-only bridge. The in-app AI session switch is off by default. MCP can inspect, stage semantic design changes, build private local candidates, and record qualification evidence. It cannot overwrite an original, save a project, export to a user destination, or distribute an output.

## Project files and encryption

- `.pstudio` is a ZIP-based self-contained package with canonical project JSON and immutable-by-hash Resources.
- `.pstudio-secure` encrypts the complete package with AES-256-GCM and PBKDF2-SHA-256.
- Presentation Studio cannot recover a lost project password.
- Project encryption does not cover the external original files or separately exported PowerPoint, PDF, SVG, or PNG files.

## Verify the source checkout

Run:

```sh
npm run quality
npm run desktop:smoke
```

A passing build proves that the current source compiles and the Electron renderer opens. It does not approve a cleaned deck for distribution; exported decks remain subject to human production, brand, content, and technical review.
