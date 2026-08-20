# Presentation Studio

Presentation Studio is a local-first Electron application for auditing, redesigning, cleaning, reviewing, and composing editable presentations. Resources are immutable inputs; one constrained semantic HTML/CSS Studio Web Scene is the current presentation shared by the human and AI. Converted ORNL Template Pack layouts and shared recipes are design vocabulary inside that scene. The exact scene revision compiles to editable PowerPoint, and Microsoft PowerPoint-native pixels and measurements remain final appearance authority.

The application never overwrites an imported deck. It embeds source bytes in a self-contained project, stages bounded changes for review, validates exact visible content against an explicit source-to-output slide map, and exports a new PowerPoint copy. Slide count may increase only through a reviewable continuation plan that preserves source order and repeats required table headers.

## Current working slice

- Existing populated ORNL title slides are sacred, source-preserved brand compositions. Studio and MCP cannot recompose, restyle, move, resize, add to, remove from, or replace them; full-deck builds transplant the exact native title-slide XML together with its approved layout, master, theme, and related media. New presentations instead use an approved ORNL title layout and edit only its intended placeholders.
- A persisted `presentation-studio/web-scene` version-5 design model that normalizes imported content onto a canonical 13.333 × 7.5-inch React/HTML/CSS canvas while retaining source-coordinate bindings and explicit editable/catalog/semantic-atom provenance; shared ORNL title/content, two-column, repeated comparison-card, objective-column, steps/evidence, challenge/evidence, process-flow, labeled-figure-grid, table, figure-grid, footer, and converted-template recipes; drag, resize, numeric geometry, bounded Aptos type controls, and a read-only source reference. Slides, comments, per-slide builds, full-deck build, and export now resolve against this one scene revision. Converted Template Pack images/vectors/fills compile with native editable text, images, vectors, and tables; exact visible-copy, source-order, table-grid, merged-cell, semantic-color, and cell-break round trips are tested.
- First-class editable table design memory. A reviewed treatment can be published as a structure-fingerprinted exemplar and applied only to compatible tables; semantic fills, exact copy, order, and merge topology remain protected. Merge-aware continuation plans repeat identified headers, split only at safe body-row boundaries, materialize as real editable PowerPoint slides during Build all, and carry explicit source/output mapping into Slides, comments, protected-title placement, and full-deck qualification.
- A hash-pinned, ignored private-golden harness qualifies six representative EMT communication jobs against source and prior design references without committing customer content. Its objective gate writes a fresh editable PPTX, rerenders and remeasures it through Microsoft PowerPoint, and requires exact text/table structure, Aptos, zero true overflow, zero off-slide objects, native table-cell clearance, and material design impact before full-size visual review.
- A complete private deck-qualification runner available as **Inspect all**, `npm run qualify:deck`, and MCP `run_deck_qualification`. It reopens the immutable source and exact central candidate in Microsoft PowerPoint, writes one 2,200-pixel PNG per slide for both, records native measurements and hashes, opens on a clean candidate overview, and provides exact full-slide comparison plus optional issue crops/diagnostic overlays. MCP can inspect paged contact sheets and record raster-bound `ready`, `revise`, or `hold` reviews. A changed build creates a new attempt with objective trend; unchanged bytes are not rerun and unresolved automatic attempt three is held. Objective passes mean ready for visual review, never better by definition.
- A governed **Concept → Editable** path: a generated or human-supplied image remains an immutable `concept-only` Resource, is exposed as a bounded preview by the single active AI access switch, and records only the composition, hierarchy, negative-space, color, figure, image-treatment, or rhythm characteristics the user/agent may follow. Generated wording, logos, data, claims, and technical details remain untrusted; Studio reconstructs the approved direction as editable objects and validates it through PowerPoint.
- A model-independent **visual-needs queue** for layout concepts, figure concepts, image treatments, supporting visuals, and diagram rebuilds. Each source-hash-bound brief defaults to abstract structure only, selects an affirmative restrained/balanced/expressive ORNL recipe, requests one text-free/logo-free concept raster, closes forbidden content, and can be fulfilled by any MCP-capable AI without embedding a provider in Studio. Human attach/detach controls use image Resources automatically preview-shared by the active AI switch; the need progresses through concept, editable reconstruction, and PowerPoint-native review instead of treating generated pixels as the slide.
- A deterministic **concept zones → editable reconstruction** bridge. `reconstruct_studio_concept` maps only the approved normalized title, visual, evidence, caption, and other zones to exact source-bound Studio nodes through a shared recipe; it never traces generated pixels or copies generated content. Studio reports material design impact and refuses to advance a visual need when the result is unchanged, typography-only, or a small cleanup instead of the requested layout, image, figure, or verified-diagram reconstruction.
- The active ORNL Brand Agent V2.3.2 prompt and 20-file Knowledge package may be mirrored into `.presentation-studio/private-brand-agent-v2.3.2` for local development reference. That directory is ignored, never packaged or committed, and informs the normalized design standard rather than becoming an automatic external-upload source.
- Constraint-driven collaborative layout in that same scene: deck rhythm tokens, PowerPoint-bound optical insets, safe/center/peer guides, grid snapping, Shift multi-select, atomic relationship-group movement, persistent optical/structural alignment, equal-gap distribution, and safe-region fitting. MCP models use `refine_studio_layout` instead of inventing correction coordinates.
- Source-locked technical-figure decisions shared by the canvas and MCP: preserve a dense figure as one evidence unit, preserve and frame it with restrained ORNL geometry, record a hybrid rebuild that retains source screenshots/data, or hold a full redraw candidate for content verification. The original remains visible until any replacement passes its information/relationship invariants and PowerPoint-native intent review.
- Object-isolated PowerPoint evidence for fresh compositions. When a source-locked technical group is reused inside a new web-designed layout, Studio builds a private one-slide render source that hides neighboring top-level shapes while retaining the complete requested PowerPoint group, then places that native raster as one meaning-preserving unit. This avoids accidental title, footer, or nearby-object leakage without tracing or flattening the rest of the slide.
- Production design recipes for question-plus-diagram, coupled-evidence, paired-evidence, challenge, process, objective, table, and comparison communication jobs. Question pages receive a deliberate prompt rail and dominant technical evidence field; paired figures receive equal neutral fields with source aspect ratios preserved instead of one-off redraws.
- First-class figure controls add a group frame, normalized crop and focal point, aspect lock, and explicit caption/label/callout/connector relationships. Complete source-locked figures can move through layout constraints without unlocking or reconstructing their internal technical content.
- Multi-file `.pptx` import with bounded OOXML preflight, unsafe-path checks, and expanded-size/compression-ratio limits.
- App-wide drag-and-drop and Resources file-picker intake for presentations, documents, data, images, audio, video, and SVG assets.
- Native greenfield composition from Resources: turning on the single AI access switch automatically shares extracted TXT, Markdown, JSON, CSV/TSV, DOCX, and XLSX derivatives plus bounded image previews. Those Resources can be planned into assertion-evidence slides, instantiated from stable converted ORNL Template Pack layouts or shared Studio recipes, and persisted directly as the source-grounded `presentation-studio/web-scene` JSON without requiring a starter PowerPoint. Every slide retains immutable Resource, derivative, and exact-excerpt hashes; the first slide uses an approved ORNL title layout and becomes protected after initial compilation.
- Versioned first-run onboarding walkthrough with keyboard navigation, deterministic spotlights, persisted completion, and an always-available Tour replay control.
- Local, SHA-256-verified POTX/PPTX installation plus a Designs workspace that inventories every native master/layout, compiles duplicate-safe semantic slots/constraints, and shows PowerPoint-native layout pixels with app-only editable-region guides; production template bytes stay outside Git.
- Atomic native-layout recomposition that reuses a byte-exact approved layout already present in an imported deck when possible, remaps compatible placeholders, places source-bound objects into approved semantic zones, and requires a PowerPoint-native Current/Proposal review before acceptance.
- A revision-bound Slides gallery that shows only the latest PowerPoint-native result for each central Studio slide revision, holds stale redesigned slides until rebuilt, opens the same Studio slide for layout/editing, and supports normalized region comments with private-note and **Submit to AI** paths. Comment entry autofocuses and stays on the selected slide, Electron exposes spelling suggestions, users can delete feedback, and MCP mutations clear only exact comment IDs explicitly addressed.
- A versioned hybrid slide scene and PowerPoint preservation envelope: stable source-bound object IDs, native geometry/z-order, semantic roles, editable/preserved/conversion/blocking fidelity states, explicit per-object operation capabilities, SHA-256-bound source slide/relationship parts, project-package persistence, Deck Audit coverage, Slides-editor fidelity labels, and content-minimized MCP scene tools. Unsupported native internals remain locked and preserved rather than reconstructed or silently flattened.
- A macOS PowerPoint-native render and measurement bridge for faithful Source/Current/Proposal/Export pixels and geometry in the Slides UI, Review UI, and MCP, with serialized temporary jobs, bounded local rasterization, cleartext cleanup, exact slide-count validation, safe recovery from PowerPoint's zero-document startup shell, explicit renderer provenance, and a visibly labeled OOXML diagnostic fallback.
- A first iterative AI design loop: revision-bound per-slide and representative-deck work orders join native raster evidence, exact locked content, scene objects, ORNL rules, findings, comments, and ranked layouts; responsive Template Pack slot contracts and `fit_scene_to_layout` compose relationship-preserving groups across shared regions with hierarchy-aware spacing and bounded scaling; guarded native target-layout reuse/cloning preserves approved PowerPoint dependencies; optical text metrics drive deterministic alignment; PowerPoint-native side-by-side comparison returns raster hashes and changed-region metrics; and the AI can reject its own failed pending draft with the visual rationale/evidence recorded. Substantial unmeasured text-frame replacement remains blocked.
- A central three-pass visual critic: `get_studio_slide_critique` returns the original and exact current export-result pixels with PowerPoint-native overflow, safe-region, optical-alignment, spacing, hierarchy, density, and figure findings; `record_studio_visual_critique` withholds ready while serious issues remain and holds unresolved pass three for human review. A ready result records its content signature, recipe/layout, and deck rhythm for consistent reuse.
- Converted Template Pack layouts use the authoritative PowerPoint-native layout render as both the Studio editor base and exported artwork base, keeping inherited master imagery, marks, fills, and strokes visually exact while Studio content remains editable above it.
- Build all is one visible background production job shared by the app and MCP. It reports preflight, editable compilation, approved ORNL templating, Microsoft PowerPoint rendering, cell-level measurement, and hard-QA phases; it can be polled or safely canceled without blocking the canvas. Started, failed, canceled, and superseded jobs never become Slides/export authority. A ready candidate still receives independent whole-file content, table, render, measurement, and export acceptance checks.
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
- Save-as-you-work recovery for human and MCP edits: rapid metadata/scene checkpoints, coalesced complete project snapshots, current/previous rotation, visible save state, and an explicit Recover action from the blank startup workspace. MCP mutations do not report success until their crash checkpoint is durable; mutations that add or change embedded Resources wait for a complete self-contained snapshot.
- Local STDIO MCP server with a loopback-only desktop bridge, one AI access switch that automatically shares all compatible embedded Resources, a complete semantic Template Pack catalog, authoritative per-layout PowerPoint images, deterministic per-slide layout recommendations, versioned AI design work orders, hybrid-scene fidelity/preservation context, proposal-only cleanup/recomposition tools, exact object-mapped content/geometry context, validated atomic 1–20 object layout transactions, native render comparison, and evidence-bound rejection of the AI's own failed pending design draft.
- One versioned machine-readable Presentation Design Standard shared by project defaults, the Rules UI, tests, and MCP: 16:9, Aptos, current authorized ORNL Template Pack, exact-content preservation, editable output, autonomous routine design choices, and deterministic native-table fallback tokens.
- A model-independent MCP Designer Contract that requires deck-wide visual improvement, exact approved-content preservation, per-text-box fit/alignment review, compatible template-layout selection, restrained semantic visuals, and independent export-render QA; authorized clients can read exact design context, the resolved cleanup profile, Current or Proposal JPEG renders, submitted location-bound design threads, and can stage the full deck-wide designer proposal.
- Visible local MCP progress that moves through **Found issues**, **Fixing**, and **Rechecking original intent**, then reports ready/attention without exposing raw prompts or presentation content. A claimed visual improvement is withheld when exact text, source visual identity, or meaning-bearing relationships remain unverified.
- One-line macOS and Windows source installers that download the release without Git, prepare a SHA-256-verified portable Node.js runtime when needed, verify the staged app, and create a reusable launcher; no unsigned packaged app installer.

The bounded layout path has completed a live end-to-end qualification on a 26-slide local deck: MCP inspection and staging, in-app Current/Proposal review, exact-content export validation, and full-deck Microsoft PowerPoint JPEG comparison. Only the intended slide render changed; all 25 unaffected renders were byte-identical. This is evidence for the current object-layout workflow, not a claim that autonomous deck-wide redesign or native render QA is complete.

Specification 1.0 makes PowerPoint-native rendering and measurement, the complete ORNL Template Pack compiler, the hybrid editable-scene/preservation model, deterministic geometry solvers, and the AI visual-design loop the architectural foundation. The macOS bridge now supplies native object/text/table-cell measurements, image-bearing inspection packets, high-resolution crops, paged deck contact sheets, responsive semantic layout solvers, bounded visual critique, and independently rerendered export acceptance. The version-2 synthetic native canary remains the qualification gate; its complete 14-slide Current/Proposal/Export run passed on macOS PowerPoint 16.111.3 on 2026-08-17. Editable table continuation is implemented; Windows native automation, general non-table continuation composition, independent Template Pack artwork/content-recipe coupling, persistent native-render caching, and the broader Canva-style editor remain. The OOXML renderer remains an honest diagnostic/editor fallback and is never presented as PowerPoint-render fidelity. Cleaned output remains a review copy until the actual exported artifact passes native comparison. The product and technical direction is in [PRESENTATION-STUDIO-SPEC.md](PRESENTATION-STUDIO-SPEC.md), with the implementation order in [TODO.md](TODO.md).

## Source setup

Requirements: macOS or Windows, an internet connection during installation, and enough local disk space. The one-line installer checks Node.js and npm and installs a verified user-local runtime when either is missing or too old. It does not require Git or administrator access under normal user permissions. The locked app includes its own PDF.js slide rasterizer, so Poppler and `pdftoppm` are not required.

Version 0.3.1 is currently delivered from the isolated `codex/web-slide-design-engine` branch.

### macOS - one line

```sh
curl -fsSL https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-macos.sh | /bin/zsh
```

This installs to `~/Applications/Presentation Studio`, verifies the staged application, creates `Launch Presentation Studio.command`, and starts the app. Run the same command again to update the managed installation.

### Windows - one line

Run in PowerShell:

```powershell
irm https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-windows.ps1 | iex
```

This installs to `%LOCALAPPDATA%\Presentation Studio`, verifies the staged application, creates `Launch Presentation Studio.cmd`, and starts the app. Run the same command again to update the managed installation.

The installers do not change Gatekeeper, PowerShell execution policy, SmartScreen, or other operating-system protections. Microsoft PowerPoint is optional for launching the app but is required for PowerPoint-native rendering and final native validation; licensed Microsoft software is not installed automatically.

### Manual Git setup

Developers who want a Git checkout can instead run this as one complete shell line:

```sh
git clone --branch codex/web-slide-design-engine --single-branch https://github.com/adammalin/Presentation-Studio.git && cd Presentation-Studio && ./scripts/setup-macos.command
```

To update an existing developer checkout without merging into `main`:

```sh
git fetch origin codex/web-slide-design-engine
git switch codex/web-slide-design-engine
git pull --ff-only origin codex/web-slide-design-engine
```

#### macOS checkout

Double-click `scripts/setup-macos.command`, or run:

```sh
./scripts/setup-macos.command
```

Then start with `scripts/start-macos.command`.

#### Windows checkout

In PowerShell from the repository:

```powershell
& .\scripts\setup-windows.ps1
& .\scripts\start-windows.ps1
```

The checkout setup scripts install locked npm dependencies, including the bundled PDF.js slide rasterizer, run tests, build the renderer, run the repository data-safety scan, and print the MCP configuration snippet. They do not install Node.js; the one-line installers above provide that bootstrap layer. The one-line installers also register the installed server automatically when a local Codex configuration is detected.

## Development

```sh
npm ci
npm test
npm run build
npm start
```

To qualify representative slides through fresh composition and Microsoft PowerPoint-native rendering without adding the source deck to the repository:

```sh
npm run benchmark:studio-web -- --source /absolute/path/source.pptx --slides 2,6,21 --output /tmp/presentation-studio-web-benchmark

# Exercise exact installed-template artwork and semantic layout selection locally.
# The authorized template, source deck, candidates, and renders remain outside Git.
npm run benchmark:studio-web -- --source /absolute/path/source.pptx --slides 2,6,21 --template /absolute/path/authorized-template.potx --design-mode template --output /tmp/presentation-studio-template-benchmark

# Private hash-pinned visual regression; version 2 can also pin the authorized
# ORNL template. It writes selected source/candidate/golden PNGs, an editable
# candidate PPTX, a JSON ledger, an HTML triptych review, and a context-isolated
# agent prompt. All private inputs and evidence remain outside Git.
npm run benchmark:private-golden -- --manifest /absolute/path/private-golden-manifest.json --output /tmp/presentation-studio-private-golden

# Full source/candidate deck qualification. The output folder must be new and
# remains private local evidence outside Git.
npm run qualify:deck -- --source /absolute/path/source.pptx --candidate /absolute/path/candidate.pptx --output /absolute/path/new-qualification-run --protected-slides 1

# Portable release checks: typecheck, tests, production build, and data-safety.
npm run quality

# Approved local/self-hosted PowerPoint workstation only: portable checks,
# desktop smoke, and the synthetic native canary. No client deck is uploaded.
npm run quality:native
```

The hosted GitHub workflow runs only the portable lane. The manual native workflow requires an explicitly provisioned self-hosted macOS runner with Microsoft PowerPoint and uses the synthetic canary; defining the workflow does not provision or authorize a runner. The full boundary and evidence contract are documented in [docs/QUALITY-PIPELINE.md](docs/QUALITY-PIPELINE.md).

To run the MCP server manually:

```sh
npm run mcp
```

Print a model-independent MCP client configuration:

```sh
node scripts/configure-mcp.mjs
```

The desktop app must be open for project-specific tools. `get_design_contract` is always available so any MCP-capable model can read the required behavior before working. Once a deck is open, `get_agent_runbook` turns the current project state into one concrete next action, intervention-level guidance, representative archetype coverage, source-wins acceptance rules, and deck-consistency requirements. Read operations beyond basic app status/design instructions and every proposal tool require the user to enable the visible AI access switch. That one switch automatically shares every embedded Resource at the highest level Studio supports: extracted text for compatible documents/data, bounded previews for images, and metadata for other formats. MCP can create a source-grounded native Studio presentation from those sources, build a private local qualification evidence bundle, read contact sheets and exact slide images, request issue crops/overlays, and record raster-bound qualification reviews. Those operations still cannot apply a legacy proposal, save a project, export a deck to a user destination, retrieve original Resource bytes, or distribute an output.

## Data boundary

Do not commit official ORNL templates or extracted template assets, client presentations, papers, manuscripts, project packages, recovery files, extracted text, qualification runs, previews, or exports. Synthetic fixtures are generated locally under `fixtures/generated/`. Imported files remain local and are copied into the user's project package by SHA-256; their original paths are never needed to reopen the project. App-created qualification evidence lives under the local Presentation Studio application-data folder and contains PNGs plus content-minimized measurement/report metadata, not another persisted PPTX copy.

Project encryption covers packaged JSON and Resource bytes only. It does not encrypt the external originals or separately exported PowerPoint, PDF, SVG, or PNG files.
