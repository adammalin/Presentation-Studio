# Presentation Studio

Presentation Studio is a local-first Electron application for auditing, redesigning, cleaning, reviewing, and composing editable presentations. Its primary design surface is now a constrained semantic HTML/CSS Studio Web Scene shared by the human and AI. Imported PowerPoint remains the exact-content preservation envelope, supported web-scene objects compile back to editable PowerPoint, and Microsoft PowerPoint-native pixels and measurements remain final appearance authority.

The application never overwrites an imported deck. It embeds source bytes in a self-contained project, stages bounded changes for review, validates that visible text and slide count remain unchanged, and exports a new PowerPoint copy.

## Current working slice

- A persisted `presentation-studio/web-scene` design model that extracts exact source-bound text, tables, pictures, and preserved native objects into semantic React/HTML/CSS slides; shared ORNL title/content, two-column, table, figure-grid, and installed-template recipes; drag, resize, numeric geometry, bounded Aptos type controls, and side-by-side source reference; MCP scene inspection and staging; and editable PowerPoint round-trip tests with exact visible-copy and table-structure preservation.
- Multi-file `.pptx` import with bounded OOXML preflight, unsafe-path checks, and expanded-size/compression-ratio limits.
- App-wide drag-and-drop and Resources file-picker intake for presentations, documents, data, images, audio, video, and SVG assets.
- Versioned first-run onboarding walkthrough with keyboard navigation, deterministic spotlights, persisted completion, and an always-available Tour replay control.
- Local, SHA-256-verified POTX/PPTX installation plus a Designs workspace that inventories every native master/layout, compiles duplicate-safe semantic slots/constraints, and shows PowerPoint-native layout pixels with app-only editable-region guides; production template bytes stay outside Git.
- Atomic native-layout recomposition that reuses a byte-exact approved layout already present in an imported deck when possible, remaps compatible placeholders, places source-bound objects into approved semantic zones, and requires a PowerPoint-native Current/Proposal review before acceptance.
- A revision-bound Slides gallery generated from each embedded editable PPTX, selectable close-up review, a first measured object editor for drag/nudge/resize/safe-area alignment, and normalized region comments with separate private-note and **Submit to AI** paths. Comment entry autofocuses and stays on the selected slide, Electron exposes spelling suggestions, users can delete feedback from Slides or Review, and MCP mutations clear only exact comment IDs explicitly addressed.
- A versioned hybrid slide scene and PowerPoint preservation envelope: stable source-bound object IDs, native geometry/z-order, semantic roles, editable/preserved/conversion/blocking fidelity states, explicit per-object operation capabilities, SHA-256-bound source slide/relationship parts, project-package persistence, Deck Audit coverage, Slides-editor fidelity labels, and content-minimized MCP scene tools. Unsupported native internals remain locked and preserved rather than reconstructed or silently flattened.
- A macOS PowerPoint-native render and measurement bridge for faithful Source/Current/Proposal/Export pixels and geometry in the Slides UI, Review UI, and MCP, with serialized temporary jobs, bounded local rasterization, cleartext cleanup, exact slide-count validation, safe recovery from PowerPoint's zero-document startup shell, explicit renderer provenance, and a visibly labeled OOXML diagnostic fallback.
- A first iterative AI design loop: revision-bound per-slide and representative-deck work orders join native raster evidence, exact locked content, scene objects, ORNL rules, findings, comments, and ranked layouts; responsive Template Pack slot contracts and `fit_scene_to_layout` compose relationship-preserving groups across shared regions with hierarchy-aware spacing and bounded scaling; guarded native target-layout reuse/cloning preserves approved PowerPoint dependencies; optical text metrics drive deterministic alignment; PowerPoint-native side-by-side comparison returns raster hashes and changed-region metrics; and the AI can reject its own failed pending draft with the visual rationale/evidence recorded. Substantial unmeasured text-frame replacement remains blocked.
- Local presentation-font registration for Aptos families supplied by an installed Microsoft PowerPoint application, with an explicit Arial/sans-serif fallback so previews never silently become serif.
- Local bounded text/value extraction for TXT, Markdown, JSON, CSV/TSV, DOCX, and XLSX Resources, with explicit stored-only or needs-review states for unprocessed formats.
- Template-family classification that requires a human decision when the exact template revision cannot be proven.
- Slide, master, layout, theme, notes, comments, media, table, chart, picture, and font inventory.
- Native table style/structure fingerprints, picture treatment/description inventory, and bounded production findings.
- Deck-wide Designer Cleanup proposals that inventory every directly editable text box, normalize supported legacy fonts and compatible native tables, repair collision-checked high-confidence cover/peer alignment drift, flag overflow/off-slide/safe-margin/ambiguous geometry, and assign explicit changed, approved-as-is, or needs-review dispositions.
- Batch table-style exemplar registration by embedded source hash, slide, and object ordinal.
- Content-minimized JSON audit reports that omit slide text, notes, picture names/descriptions, and Resource bytes.
- Source-preserving PPTX export with slide-count, visible-text, exact table-cell-content, merged-table-topology, package-wide XML syntax, independent PowerPoint measurement, and native raster validation.
- Self-contained `.pstudio` packages with original Resource bytes, local derivatives, and independent integrity hashes.
- Optional AES-256-GCM encrypted `.pstudio-secure` packages using PBKDF2-SHA-256.
- Local STDIO MCP server with a loopback-only desktop bridge, session disclosure switch, explicitly authorized Resource metadata, a complete semantic Template Pack catalog, authoritative per-layout PowerPoint images, deterministic per-slide layout recommendations, versioned AI design work orders, hybrid-scene fidelity/preservation context, proposal-only cleanup/recomposition tools, exact object-mapped content/geometry context, validated atomic 1–20 object layout transactions, native render comparison, and evidence-bound rejection of the AI's own failed pending design draft.
- One versioned machine-readable Presentation Design Standard shared by project defaults, the Rules UI, tests, and MCP: 16:9, Aptos, current authorized ORNL Template Pack, exact-content preservation, editable output, autonomous routine design choices, and deterministic native-table fallback tokens.
- A model-independent MCP Designer Contract that requires deck-wide visual improvement, exact approved-content preservation, per-text-box fit/alignment review, compatible template-layout selection, restrained semantic visuals, and independent export-render QA; authorized clients can read exact design context, the resolved cleanup profile, Current or Proposal JPEG renders, submitted location-bound design threads, and can stage the full deck-wide designer proposal.
- Visible local MCP activity/completion/failure feedback that identifies the operation without exposing raw prompts or presentation content.
- Source-based macOS and Windows setup/start scripts; no app installer.

The bounded layout path has completed a live end-to-end qualification on a 26-slide local deck: MCP inspection and staging, in-app Current/Proposal review, exact-content export validation, and full-deck Microsoft PowerPoint JPEG comparison. Only the intended slide render changed; all 25 unaffected renders were byte-identical. This is evidence for the current object-layout workflow, not a claim that autonomous deck-wide redesign or native render QA is complete.

Specification 1.0 makes PowerPoint-native rendering and measurement, the complete ORNL Template Pack compiler, the hybrid editable-scene/preservation model, deterministic geometry solvers, and the AI visual-design loop the architectural foundation. The macOS bridge now supplies native object/text/table-cell measurements, image-bearing inspection packets, high-resolution crops, paged deck contact sheets, responsive semantic layout solvers, bounded visual critique, and independently rerendered export acceptance. The version-2 synthetic native canary remains the qualification gate; its complete 14-slide Current/Proposal/Export run passed on macOS PowerPoint 16.111.2 on 2026-08-13. Windows automation, continuation-slide orchestration, cloned-master artwork qualification, persistent native-render caching, and the broader Canva-style editor remain. The OOXML renderer remains an honest diagnostic/editor fallback and is never presented as PowerPoint-render fidelity. Cleaned output remains a review copy until the actual exported artifact passes native comparison. The product and technical direction is in [PRESENTATION-STUDIO-SPEC.md](PRESENTATION-STUDIO-SPEC.md), with the implementation order in [TODO.md](TODO.md).

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
