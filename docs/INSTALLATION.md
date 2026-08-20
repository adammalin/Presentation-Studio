# Presentation Studio source installation

Presentation Studio is distributed through a one-line source installer during the initial release. There is no DMG, PKG, MSI, EXE, app-store package, or other unsigned packaged application installer.

This guide installs Presentation Studio 0.3.1 from the isolated `codex/web-slide-design-engine` branch. It does not merge or change the repository's `main` branch.

## Requirements

- macOS or Windows
- Internet access during installation
- Local disk space for source decks, self-contained project packages, and exported copies

The installer checks Node.js and npm. If Node.js 22.13 or newer is not available, it downloads the official Node.js 22.13 portable runtime, verifies its SHA-256 checksum against the official manifest, and keeps it inside the Presentation Studio install folder. Git is not required.

Microsoft PowerPoint is optional for launching Presentation Studio but required for PowerPoint-native rendering and final native validation. The locked application dependencies include a local PDF.js rasterizer, so Poppler, `pdftoppm`, and Homebrew are not required. The installer cannot install or license Microsoft Office.

## macOS - one line

Paste this entire line into Terminal:

```sh
curl -fsSL https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-macos.sh | /bin/zsh
```

The managed installation is placed in `~/Applications/Presentation Studio`. The installer builds and tests a staged copy before replacing an earlier managed version, creates `Launch Presentation Studio.command`, and starts the app. Run the same one-line command again to update it. For an update, the installer closes only the verified managed Presentation Studio process before activating and reopening the new version so the app and MCP cannot remain on different design-standard versions.

The script does not alter Gatekeeper, quarantine settings, or other operating-system protections. Follow your organization's approved software process if policy blocks execution.

## Windows - one line

Paste this entire line into PowerShell:

```powershell
irm https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-windows.ps1 | iex
```

The managed installation is placed in `%LOCALAPPDATA%\Presentation Studio`. The installer builds and tests a staged copy before replacing an earlier managed version, creates `Launch Presentation Studio.cmd`, and starts the app. Run the same one-line command again to update it. For an update, the installer closes only the verified managed Presentation Studio process before activating and reopening the new version so the app and MCP cannot remain on different design-standard versions.

The script does not alter PowerShell execution policy, SmartScreen, or other operating-system protections. Follow your organization's approved software process if policy blocks execution.

## Manual Git checkout for developers

The one-line installers do not require Git. Developers who want a source checkout can use one complete command on macOS:

```sh
git clone --branch codex/web-slide-design-engine --single-branch https://github.com/adammalin/Presentation-Studio.git && cd Presentation-Studio && ./scripts/setup-macos.command
```

To update an existing Git checkout, close Presentation Studio and run:

```sh
git fetch origin codex/web-slide-design-engine
git switch codex/web-slide-design-engine
git pull --ff-only origin codex/web-slide-design-engine
```

The fast-forward-only pull stops instead of overwriting local development changes. After updating, run the platform setup script so dependencies and the production renderer match the checked-out version.

### macOS developer checkout

1. Get or update the release branch using the commands above.
2. Double-click `scripts/setup-macos.command`, or run it from Terminal in the repository.
3. The script runs `npm ci`, tests, the repository data-safety scan, and a production build.
4. Start Presentation Studio with `scripts/start-macos.command`.

The script does not change Gatekeeper or other operating-system protections.

### Windows developer checkout

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

Presentation Studio must be open. The MCP server uses STDIO and connects to the active app through a per-session token on a loopback-only bridge. The one-line installer automatically registers the installed server when it detects Codex; restart Codex afterward so it reloads the server list. Other MCP clients can use the printed standard configuration entry. The in-app AI access switch is off by default. MCP can inspect, stage semantic design changes, build private local candidates, and record qualification evidence. It cannot overwrite an original, save a project, export to a user destination, or distribute an output.

## Start a ChatGPT Desktop design session

Use this after Presentation Studio is open, the MCP connection is enabled, and the approved project material is in place:

```text
Connect to the Presentation Studio MCP and work in the currently open project.

First, read the Presentation Studio design contract, check the app status, then call get_agent_runbook and follow its current next action. Inventory the authorized project Resources and inspect the installed ORNL Template Pack. Confirm that you can read the source content - not merely filenames or metadata. If anything required is inaccessible, tell me exactly what must be shared or attached. Do not invent missing content.

For a source-only project, read every required compatible text Resource completely with get_resource_text, inspect stable layout IDs with get_template_layout_catalog, then call create_studio_presentation once with the complete source-grounded deck plan. Create it without a starter PowerPoint; do not invent one.

Create a polished, editable, 16:9 ORNL presentation from the supplied source materials.

Content direction:
- Organize the material into a clear narrative using assertion-evidence slides.
- Give each slide one primary takeaway and supporting evidence.
- You may condense source prose, but preserve technical meaning, names, numbers, units, qualifications, and attribution.
- Preserve approved or locked copy exactly. Do not introduce unsupported claims, data, diagrams, or conclusions.
- Infer routine structure and design choices. Ask only about genuine audience, technical, content-authority, or approval ambiguities.

Design direction:
- Use Aptos and the current approved ORNL Template Pack.
- For a new title slide, use an approved ORNL title layout and edit only intended placeholders. Never alter its artwork, marks, master, or layout.
- Make substantive whole-slide composition decisions using shared Studio recipes and compatible ORNL layouts. Do not merely keep the source arrangement or make text smaller.
- Establish one deck-wide system for titles, spacing, alignment, figures, captions, tables, colors, and repeated components.
- Use authorized Resource images only when they support the message. Preserve technical figures as relationship-aware groups unless a verified editable reconstruction is clearer.
- Keep tables editable and readable; preserve meaning-bearing colors.

Workflow:
1. Develop the narrative and slide plan.
2. Create the presentation in the single central Studio HTML/CSS scene.
3. Start the complete editable PowerPoint build, then poll get_studio_presentation_build_status until that exact background job reports ready. Do not treat started, measuring, failed, canceled, or superseded as ready.
4. Inspect the PowerPoint-native contact sheet and every full-size candidate slide.
5. Run Found issues -> Fixing -> Rechecking original intent. For every slide, compare source strengths, candidate improvements, and candidate regressions. If the candidate is weaker within its preserve, polish, recompose, or rebuild-figure intervention level, the source wins.
6. Correct overflow, alignment, hierarchy, spacing, tables, missing imagery, deck-system inconsistencies, and message drift. Qualify representative slides from every communication archetype before scaling a pattern across the deck.
7. Do not call the presentation ready while any blocker or major visual issue remains.

Use your best design judgment and minimize routine questions. Leave the completed central design visible for my review. Do not save or export the final PowerPoint until I explicitly request it.
```

Resource access is automatic while the single **AI access** switch is on. That single switch shares every embedded Resource at the highest level Studio supports: bounded extracted text for compatible documents/data, bounded previews for images, and metadata for other formats. Turning it off removes all project Resource access at once. Presentation Studio 0.3.1 can create a brand-new native Studio JSON deck directly from those sources and the installed ORNL Template Pack. PDF text extraction, legacy DOC/XLS extraction, and raw original-file retrieval are not available; convert those sources to a supported format or provide another approved source rather than inventing content. Use only material approved for the selected AI environment.

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
