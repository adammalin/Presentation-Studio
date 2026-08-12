# Presentation Studio

Presentation Studio is a local-first Electron application for auditing, conservatively cleaning, reviewing, and eventually composing editable presentations. The first working slice targets high-volume PowerPoint production cleanup: mixed fonts, template uncertainty, table/figure consistency review, unsupported package features, and human-approved output.

The application never overwrites an imported deck. It embeds source bytes in a self-contained project, stages bounded changes for review, validates that visible text and slide count remain unchanged, and exports a new PowerPoint copy.

## Current working slice

- Multi-file `.pptx` import with bounded OOXML preflight, unsafe-path checks, and expanded-size/compression-ratio limits.
- App-wide drag-and-drop and Resources file-picker intake for presentations, documents, data, images, audio, video, and SVG assets.
- Versioned first-run onboarding walkthrough with keyboard navigation, deterministic spotlights, persisted completion, and an always-available Tour replay control.
- Local, SHA-256-verified POTX/PPTX installation plus a Designs workspace that inventories and previews every native master/layout with actual template media and app-only placeholder guides; production template bytes stay outside Git.
- A revision-bound Slides gallery generated from each embedded editable PPTX, selectable close-up review, and normalized region comments with separate private-note and **Submit to AI** paths.
- Local presentation-font registration for Aptos families supplied by an installed Microsoft PowerPoint application, with an explicit Arial/sans-serif fallback so previews never silently become serif.
- Local bounded text/value extraction for TXT, Markdown, JSON, CSV/TSV, DOCX, and XLSX Resources, with explicit stored-only or needs-review states for unprocessed formats.
- Template-family classification that requires a human decision when the exact template revision cannot be proven.
- Slide, master, layout, theme, notes, comments, media, table, chart, picture, and font inventory.
- Native table style/structure fingerprints, picture treatment/description inventory, and bounded production findings.
- Deck-wide Designer Cleanup proposals that review every slide, normalize supported legacy fonts and compatible native tables, repair high-confidence cover-text alignment drift, and assign explicit changed, approved-as-is, or needs-review dispositions.
- Batch table-style exemplar registration by embedded source hash, slide, and object ordinal.
- Content-minimized JSON audit reports that omit slide text, notes, picture names/descriptions, and Resource bytes.
- Source-preserving PPTX export with slide-count, visible-text, exact table-cell-content, and merged-table-topology validation.
- Self-contained `.pstudio` packages with original Resource bytes, local derivatives, and independent integrity hashes.
- Optional AES-256-GCM encrypted `.pstudio-secure` packages using PBKDF2-SHA-256.
- Local STDIO MCP server with a loopback-only desktop bridge, session disclosure switch, explicitly authorized Resource metadata, and proposal-only write tools.
- One versioned machine-readable Presentation Design Standard shared by project defaults, the Rules UI, tests, and MCP: 16:9, Aptos, current authorized ORNL Template Pack, exact-content preservation, editable output, autonomous routine design choices, and deterministic native-table fallback tokens.
- A model-independent MCP Designer Contract that requires deck-wide visual improvement, exact approved-content preservation, per-text-box fit/alignment review, compatible template-layout selection, restrained semantic visuals, and independent export-render QA; authorized clients can read exact design context, the resolved cleanup profile, Current or Proposal JPEG renders, submitted location-bound design threads, and can stage the full deck-wide designer proposal.
- Visible local MCP activity/completion/failure feedback that identifies the operation without exposing raw prompts or presentation content.
- Source-based macOS and Windows setup/start scripts; no app installer.

PowerPoint-native slide rendering, the Canva-style direct-manipulation canvas, object/cell-stable comment anchors, the first-class native table editor and layout solver, full exemplar-driven normalization, complete semantic Template Pack compilation, PDF/SVG/PNG presentation export, and assertion-evidence composition are intentionally later slices. The Current/Proposal renderer is an honest local OOXML reconstruction of editable text, shapes, media, connectors, and native tables—not a claim of PowerPoint-render fidelity. Until independent native export rendering is available, cleaned PowerPoint output is explicitly a review copy and remains in `Needs manual review`. The product and technical direction is in [PRESENTATION-STUDIO-SPEC.md](PRESENTATION-STUDIO-SPEC.md), with the next implementation units in [TODO.md](TODO.md).

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

The desktop app must be open for project-specific tools. `get_design_contract` is always available so any MCP-capable model can read the required behavior before working. Read operations beyond basic app status/design instructions and every proposal tool require the user to enable the visible AI session switch. Resource metadata also requires a per-Resource session choice. MCP cannot apply a proposal, save a project, export a deck, or retrieve imported Resource bytes or extracted document text.

## Data boundary

Do not commit official ORNL templates or extracted template assets, client presentations, papers, manuscripts, project packages, recovery files, extracted text, previews, or exports. Synthetic fixtures are generated locally under `fixtures/generated/`. Imported files remain local and are copied into the user's project package by SHA-256; their original paths are never needed to reopen the project.

Project encryption covers packaged JSON and Resource bytes only. It does not encrypt the external originals or separately exported PowerPoint, PDF, SVG, or PNG files.
