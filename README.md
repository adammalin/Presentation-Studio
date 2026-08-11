# Presentation Studio

Presentation Studio is a local-first Electron application for auditing, conservatively cleaning, reviewing, and eventually composing editable presentations. The first working slice targets high-volume PowerPoint production cleanup: mixed fonts, template uncertainty, table/figure consistency review, unsupported package features, and human-approved output.

The application never overwrites an imported deck. It embeds source bytes in a self-contained project, stages bounded changes for review, validates that visible text and slide count remain unchanged, and exports a new PowerPoint copy.

## Current working slice

- Multi-file `.pptx` import with bounded OOXML preflight, unsafe-path checks, and expanded-size/compression-ratio limits.
- App-wide drag-and-drop and Resources file-picker intake for presentations, documents, data, images, audio, video, and SVG assets.
- Versioned first-run onboarding walkthrough with keyboard navigation, deterministic spotlights, persisted completion, and an always-available Tour replay control.
- Local bounded text/value extraction for TXT, Markdown, JSON, CSV/TSV, DOCX, and XLSX Resources, with explicit stored-only or needs-review states for unprocessed formats.
- Template-family classification that requires a human decision when the exact template revision cannot be proven.
- Slide, master, layout, theme, notes, comments, media, table, chart, picture, and font inventory.
- Native table style/structure fingerprints, picture treatment/description inventory, and bounded production findings.
- Conservative Century Gothic/Arial-to-Aptos cleanup proposals for a human-confirmed ORNL target.
- Batch table-style exemplar registration by embedded source hash, slide, and object ordinal.
- Content-minimized JSON audit reports that omit slide text, notes, picture names/descriptions, and Resource bytes.
- Source-preserving PPTX export with before/after visible-text hashes and slide-count validation.
- Self-contained `.pstudio` packages with original Resource bytes, local derivatives, and independent integrity hashes.
- Optional AES-256-GCM encrypted `.pstudio-secure` packages using PBKDF2-SHA-256.
- Local STDIO MCP server with a loopback-only desktop bridge, session disclosure switch, explicitly authorized Resource metadata, and proposal-only write tools.
- Source-based macOS and Windows setup/start scripts; no app installer.

Visual slide rendering, exemplar-driven table normalization, full template-pack compilation, PDF/SVG/PNG presentation export, and assertion-evidence composition are intentionally later slices. Until native before/after rendering is available, cleaned PowerPoint output is explicitly a review copy and remains in `Needs manual review`. The product and technical direction is in [PRESENTATION-STUDIO-SPEC.md](PRESENTATION-STUDIO-SPEC.md).

## Source setup

Requirements: Node.js 22.13 or newer, npm, and macOS or Windows.

### macOS

Double-click `scripts/setup-macos.command`, or run:

```sh
./scripts/setup-macos.command
```

Then start with `scripts/start-macos.command`.

### Windows

In PowerShell from the repository:

```powershell
& .\scripts\setup-windows.ps1
& .\scripts\start-windows.ps1
```

The setup scripts install locked npm dependencies, run tests, build the renderer, run the repository data-safety scan, and print the MCP configuration snippet. They do not disable operating-system protections and do not create an installer.

## Development

```sh
npm ci
npm test
npm run build
npm start
```

To run the MCP server manually:

```sh
npm run mcp
```

Print a model-independent MCP client configuration:

```sh
node scripts/configure-mcp.mjs
```

The desktop app must be open. Read operations beyond basic app status and every proposal tool require the user to enable the visible AI session switch. Resource metadata also requires a per-Resource session choice. MCP cannot apply a proposal, save a project, export a deck, or retrieve imported Resource bytes or extracted document text.

## Data boundary

Do not commit client presentations, papers, manuscripts, project packages, recovery files, extracted text, previews, or exports. Synthetic fixtures are generated locally under `fixtures/generated/`. Imported files remain local and are copied into the user's project package by SHA-256; their original paths are never needed to reopen the project.

Project encryption covers packaged JSON and Resource bytes only. It does not encrypt the external originals or separately exported PowerPoint, PDF, SVG, or PNG files.
