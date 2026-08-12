# Presentation Studio

## Product and Technical Specification

- **Status:** Working foundation; phased implementation in progress
- **Specification version:** 0.9
- **Date:** 2026-08-12
- **Repository:** `adammalin/Presentation-Studio`
- **Distribution boundary:** Source-based installation only; no application installers

Presentation Studio is a local-first Electron desktop application for auditing, conservatively cleaning, reflowing, composing, reviewing, and exporting presentation decks from a structured, self-contained project package. Its first production job is to help a human review and consistently clean a large batch of existing PowerPoint slides without rewriting content, erasing intentional prior edits, heavily redesigning the material, or applying the wrong organization’s template. The package contains a canonical JSON presentation model for supported editable properties, an immutable PowerPoint preservation envelope for native content that is not fully interpreted, and the documents, data, images, audio, video, and other project Resources required to reopen the work without relying on the files' original locations. The application uses versioned presentation rules and Template Packs, a PowerPoint-aware scene graph, native-render observations, and a local Model Context Protocol (MCP) server so any MCP-capable AI client can inspect an authorized project, design against reliable visual evidence, and stage changes for visible human review.

The first template target is the official ORNL 16:9 PowerPoint template. The application must preserve its master/layout intent, Aptos typography, theme colors, and presentation guidance while treating every cleaned, reflowed, or generated output as a draft until the appropriate content owner and brand reviewer approve it.

This document defines the intended product, data model, user experience, local security boundary, MCP contract, export requirements, source-based setup, installation PDF, validation gates, and phased implementation plan. The repository now contains a tested cleanup-first working foundation. Capabilities not listed in the implementation snapshot remain requirements rather than completed features.

### 0.1 Implementation snapshot - 2026-08-12

Implemented and locally qualified:

- ORNL-aligned Electron, React, TypeScript, and Vite application shell with Batch, Deck Audit, Slides, Designs, Rules, Review, and Resources workspaces;
- app-wide drag-and-drop plus multi-file Resources intake with local type/signature checks, SHA-256 deduplication, explicit processing states, and project-size limits;
- versioned first-run onboarding with deterministic target spotlights, keyboard navigation, replay from the top bar, and local completion preferences outside project files;
- local POTX/PPTX Template Pack installation with independent SHA-256 verification, protected application-data storage, restart persistence, duplicate-safe semantic slot/capacity/constraint compilation, deterministic content-compatibility scoring, and a searchable Designs gallery whose base layouts are materialized and rendered through Microsoft PowerPoint with Studio-only semantic guides;
- packaged extracted-text derivatives for TXT, Markdown, JSON, CSV/TSV, DOCX, and XLSX Resources, with explicit stored-only or needs-review handling for other accepted types;
- multi-file PowerPoint intake through immutable embedded Resources and bounded raw-OOXML preflight;
- slide, master, layout, theme, notes, modern-comment, media, chart, direct-versus-template font, native-table style/structure fingerprint, and picture treatment/description inventory;
- conservative template classification that requires a human target decision when an exact authorized template hash is unavailable;
- deck-wide Designer Cleanup proposals with per-slide dispositions, source/proposal OOXML renders, source-preserving Arial/Century Gothic-to-Aptos transformation, directly editable text-box geometry/fit inventory, collision-checked high-confidence cover and peer alignment repair, explicit overflow/off-slide/safe-margin/ambiguous-alignment review items, deterministic normalization of compatible native tables, and visible dense/complex/semantic-color exceptions;
- stable slide-local IDs, exact object-mapped text, and EMU/inch geometry for supported text, shape, picture, table, chart, connector, group, and graphic-frame objects, plus a first non-destructive Edit mode for selection, snapped dragging, measured move/resize, nudge, safe-area alignment, and shared human/MCP geometry proposals with Current/Proposal review;
- atomic MCP layout transactions covering 1–20 objects on one slide, with stale-write guards and pre-proposal rejection of new overlap, worsened text fit, unsafe-edge regression, off-slide geometry, or accidental picture distortion unless a deliberate exception is recorded;
- slide-count, visible-text-hash, exact table-cell-content, and merged-topology export guards, ZIP integrity qualification, new-copy-only export, and mandatory manual visual-review status;
- approved native-table exemplar registration by embedded Resource hash, slide number, and table ordinal;
- content-minimized JSON audit reports that omit slide text, notes, picture names/descriptions, and Resource bytes;
- self-contained `.pstudio` package Save/Open with independent original/derivative hash validation and no required original paths;
- optional full-package `.pstudio-secure` AES-256-GCM encryption with PBKDF2-SHA-256, wrong-password/tamper rejection, and encrypted recovery behavior;
- loopback-only, per-session desktop bridge and standard STDIO MCP server with status, authorized audit/Resource metadata, full semantic Template Pack catalog, deterministic per-slide layout recommendations, stable object inventory, exact bounded design context, Current/Proposal slide renders, and proposal-only font, deck-wide designer-cleanup, single-object geometry, or atomic multi-object layout tools;
- a first macOS PowerPoint-native render bridge that materializes exact temporary Source/Current/Proposal/Export PPTX revisions, exports through Microsoft PowerPoint, rasterizes bounded slide images locally, validates slide-image count, removes the cleartext job, and supplies explicit authority/provenance to Slides, Review, and MCP while retaining a visibly labeled OOXML fallback;
- revision-bound AI slide-design work orders and a representative five-slide qualification set combining native Current raster identity/provenance, exact protected content, hybrid scene objects/operations, deterministic findings, submitted design comments, ORNL design rules, and ranked semantic Template Pack candidates;
- a first semantic recomposition path that binds source objects to approved layout zones while retaining the source native master/layout, blocks substantial unmeasured text-frame replacement, materializes an exact-content proposal, compares Current/Proposal through Microsoft PowerPoint with raster hashes and changed-region metrics, and lets the AI reject only its own failed pending draft with evidence and rationale recorded;
- source setup/start scripts for macOS and Windows, repository data-safety scanning, a rendered installation-guide PDF, automated tests, production build, and Electron smoke capture.

End-to-end qualification evidence on 2026-08-12 used the live Electron app and its STDIO MCP server against an authorized 26-slide local deck: the model read exact object context, staged a bounded cover alignment, visually compared Current/Proposal, accepted it in the UI, exported a new editable PPTX, and rendered source/output decks independently through Microsoft PowerPoint to 26 JPEGs each. Slide count, every visible-text hash, every editable text-box hash, table-cell hashes, and merged-table topology matched; exactly the intended first-slide native render changed, while all 25 unaffected slide JPEGs remained byte-identical. This proves the current bounded geometry path, not full autonomous deck redesign or the complete in-app native export-acceptance gate.

Important current limits:

- the current authorized ORNL reference has been installed and its one master, 30 layouts, nine preview media assets, semantic slot/constraint catalog, deterministic compatibility ranking, 30-slide PowerPoint-native layout render, and expected SHA-256 have been qualified locally on macOS; this does not yet provide side-by-side pack management, project-pinned pack snapshots, representative synthetic content filling/overflow qualification, Windows qualification, or automatic `current-ornl` classification against the active pack;
- PDF text extraction, rich Resource previews, media derivatives, paste intake, and extracted-text MCP reads are not implemented yet; those Resources remain embedded with an honest support state;
- imported-slide Source/Current/Proposal/Export views and Template Design previews prefer PowerPoint-native pixels on supported macOS systems and visibly identify the OOXML reconstruction when native rendering is unavailable; layout preview renders are cached only for the active session, selected-object crops/contact sheets, persistent cache policy, Windows automation, and an automatic saved-export rerender comparison remain, and text fit is still a conservative deterministic estimate rather than final Office font measurement; a cleaned export therefore remains `Needs manual review`;
- semantic recomposition now stages one atomic reversible transaction containing the approved native layout remap plus validated semantic object bindings. The layout path reuses an exact SHA-256-matching approved layout already present in the deck, or clones its allowlisted master/layout/theme/media graph when no exact part exists; compatible placeholder identities are remapped before only the selected slide is repointed. Text work orders expose direct text insets, paragraph margins/indents, bullet counts, and estimated optical text starts. Full inherited-style resolution, cloned-master visual qualification across every layout, and bounded autonomous revise/retry still remain;
- table exemplar style extraction/application, cell-level overflow solving, figure normalization, prior-revision comparison, protected/excluded object scopes, and cleanup reports are not implemented;
- encrypted packages and autosave currently use bounded all-in-memory processing rather than the future streaming container;
- resumable 200-slide queueing, multi-select/guides/cropping/undo and the rest of the direct-manipulation canvas, stable text/table-cell comment anchors, the first-class table editor, PDF/SVG/PNG deck export, layout-remapping reflow/manual composition, and source-grounded assertion-evidence creation remain later phases;
- macOS source setup, the real supplied deck structures, real-package encryption, PowerPoint-open qualification, and MCP gating were exercised locally; a fresh Windows checkout has not yet been qualified.

### 0.2 Specification 0.9 architecture decision

Specification 0.9 makes AI design quality and PowerPoint fidelity one architecture rather than two later validation concerns. Presentation Studio uses a hybrid fidelity model:

1. the canonical scene is authoritative for supported editable properties and semantic design intent;
2. immutable source OOXML and retained relationships are authoritative for preserved native content that the scene does not fully represent;
3. a Microsoft PowerPoint render is the preferred visual authority for source, current, proposal, and export review when PowerPoint is available;
4. a qualified alternate renderer is a labeled fallback, never an undisclosed substitute for PowerPoint fidelity;
5. the user and AI work from the same revision-bound pixels, structured objects, template constraints, content protections, and deterministic findings.

This is not a promise that every PowerPoint feature becomes fully editable in a web canvas. It is a promise that supported objects are honestly editable, unsupported objects are preserved or explicitly converted, renderer provenance is visible, and no design is called ready until the actual exported artifact has been inspected.

---

## 1. Product vision

Presentation Studio should give staff a focused production environment that feels like ORNL OrgChart Studio and USA Map Studio while solving a different problem: reliably cleaning and quality-checking existing decks at scale, then supporting source-grounded creation when a new deck is needed.

The application should support five related jobs, in this priority order:

1. **Audit and conservative cleanup:** inspect an existing deck, preserve its content and intentional work, and normalize provable template, font, table, figure, and production inconsistencies.
2. **Batch review:** triage and process multiple decks totaling hundreds of slides while keeping each source, target template, proposal history, and export separate.
3. **Template-aware reflow:** fit a deck to its correct approved template and improve hierarchy, alignment, spacing, and layout without silently rewriting or deleting locked content.
4. **Compose from source:** transform an approved paper, document, manuscript, or structured brief into an assertion-evidence presentation with traceable source support.
5. **Hybrid and manual revision:** preserve selected content exactly while allowing specifically authorized generative fields or direct human edits.

The application is not a replacement for PowerPoint or a promise of automatic technical approval. It is a template-aware presentation production and review system whose outputs remain useful and substantially editable in PowerPoint.

### 1.1 Initial client-derived product goal

The initial goal is deliberately narrower than “redesign PowerPoint with AI”:

> Given a batch approaching 200 slides across ORNL and sponsor-template decks, produce new editable PowerPoint copies that retain all source content and purposeful prior formatting, follow the correct deck-specific template, normalize font and table/figure treatment where authorized, run deterministic production checks, surface editorial or technical-review questions, and make no heavy design or wording changes unless separately requested.

The motivating conditions are:

- some slides already use the current ORNL template while pasted legacy content introduces Century Gothic, Arial, or other local formatting;
- some slides were previously cleaned by a technical editor and later changed by customers;
- a supplied table treatment may be the approved style exemplar for a deck or batch;
- some decks use sponsor templates and must not be converted to ORNL styling merely because ORNL is operating the tool;
- the slide volume and turnaround make full manual inspection costly, but human judgment remains necessary for ambiguous formatting and technical findings.

The client correspondence and source decks are private working inputs and are not copied into this public specification or repository. Synthetic fixtures reproduce only the technical structures needed for testing.

### 1.2 Secondary product goal

Once the cleanup workflow is trustworthy, the same project model, Template Packs, Resources, renderer, MCP server, and review system should support creating new source-grounded assertion-evidence presentations. Composition is a valuable shared-platform capability, but it must not delay the first conservative-cleanup vertical slice.

---

## 2. Product principles

### 2.1 Hybrid fidelity authority

Presentation Studio does not treat a simplified web reconstruction as the complete truth for an imported PowerPoint file. Authority is divided deliberately:

- the canonical JSON scene is authoritative for every supported editable property, semantic role, Template Pack binding, design rule, comment anchor, and revision transaction;
- immutable source OOXML, relationships, and retained package parts are authoritative for native content that Presentation Studio preserves but does not fully interpret;
- immutable-by-hash Resource bytes and the project-pinned Template Pack snapshot are authoritative for packaged media and template dependencies;
- Microsoft PowerPoint-native renders are the preferred visual authority for imported, current, proposed, and exported slides when PowerPoint is available;
- deterministic validators remain authoritative for exact-content, geometry, package integrity, support-state, and policy claims.

SVG, PNG, PDF, approximate Electron previews, and alternate-office renders are derived observations. They must record their renderer and version and must not silently override native pixels or preserved package data.

### 2.1.1 Object fidelity states

Every imported object and every proposed conversion has one visible support state:

| State | Contract |
| --- | --- |
| `editable-native` | Fully represented in the scene and exported as a supported editable PowerPoint object. |
| `preserved-native` | Retained losslessly in the PowerPoint preservation envelope. The app may expose safe whole-object operations, but unsupported internal properties remain locked. |
| `conversion-required` | Editing requires a disclosed conversion that may change editability or behavior. The app shows the expected loss and does not convert without explicit approval. |
| `unsupported-blocking` | Neither safe editing nor qualified passthrough is available; the slide remains visible from its native baseline and is deferred for specialist handling. |

No renderer placeholder, raster fallback, or reconstructed approximation may be presented as an editable native object.

The first implemented scene schema is `presentation-studio/scene` version 1. It binds every directly inventoried object to a stable slide/shape locator, source z-order, native geometry, semantic role, content hash where available, represented-property states, and a current-operation matrix. Its companion `presentation-studio/preservation-envelope` version 1 binds the embedded source package and each discovered slide/relationship part by SHA-256, records advanced package blockers, and declares surgical OOXML overlay as the cleanup export strategy. These schemas do not make an unsupported PowerPoint feature editable: source bytes remain authoritative for preserved content and Microsoft PowerPoint-native pixels remain authoritative for appearance.

### 2.2 Template fidelity before visual imitation

Presentation Studio must compile and preserve template structure rather than recreate a template from screenshots or approximate it from colors and fonts alone. The master, layouts, placeholders, theme, assets, and design rules are distinct parts of the template contract.

### 2.2.1 Current authorized ORNL template and complete layout catalog

Presentation Studio must keep the newest **locally installed and authorized** ORNL PowerPoint template available as a versioned Template Pack. “Latest” means the newest revision supplied or approved for the workstation or project; it does not mean an unverified Internet download, and the application must never replace or migrate a template silently.

The active ORNL Template Pack must expose every real master and approved layout/design in the installed template as a selectable option. Layout previews must be rendered from that installed template, not recreated as screenshots, CSS approximations, or visually similar substitutes. Reflow and composition must clone or use the real source master/layout, fill its semantic placeholders, and preserve inherited theme styles, Aptos typography, template artwork, footers, and editable PowerPoint structure.

A user or MCP client may request any installed ORNL layout for authorized content, but “any content into any design” does not permit content loss or forced fit. Presentation Studio must validate structural compatibility, disclose unmapped content and overflow, and stage a more suitable layout or slide split when the selected design cannot safely contain the exact content. Under `locked` or `reflow-only` policy, it must not rewrite, delete, rasterize, or silently shrink content to force a match.

### 2.3 Human-reviewed AI assistance

MCP write tools stage proposals. They do not directly overwrite the working project or saved project file. A person reviews Before and After, warnings, content-policy effects, and source traceability before choosing Apply or Reject.

Human canvas edits are different: an authorized direct manipulation is applied immediately to the working scene as one validated, undoable transaction. Human and AI edits must use the same command and scene-revision model so a drag, an alignment command, a table resize, and an AI-staged reflow cannot produce two incompatible representations of a slide. A user may also submit a location-anchored design thread as a scoped instruction to the AI; doing so authorizes analysis and a proposal for that scope, not silent application, save, export, publication, or content rewriting.

### 2.4 Content authorization is separate from model compatibility

An AI model being able to use MCP does not authorize it to receive a project Resource. Resources remain local by default. The application must disclose exactly which metadata, extracted text, data, transcript, or visual preview will enter an AI conversation and require explicit session permission before MCP returns that content.

### 2.5 Exact-content protection is enforceable

“Do not rewrite” is a deterministic content policy, not a prompt preference. Locked strings, table cells, chart data, notes, and asset identities must be compared before and after a proposed reflow. Unexpected changes fail validation.

### 2.6 PowerPoint-safe design vocabulary

The Electron canvas may use CSS for the application shell and controls, but slide effects offered as editable features must have a defined PowerPoint representation. Unsupported browser-only effects must not appear in an editable slide preview as though they will survive export.

### 2.7 Evidence from actual artifacts

Validation must inspect the generated JSON, PPTX, PDF, SVG, and PNG artifacts. A clean Electron preview alone is not sufficient proof of export fidelity. The normal AI design loop materializes a temporary candidate PPTX, renders it through PowerPoint when available, compares that native observation with the source/current/proposal evidence, and returns any divergence to the same revision before the slide may be marked ready.

### 2.8 Reversible, resumable work

The application must autosave locally, preserve undo/redo, reject stale proposals, and process large decks in cancelable, resumable batches.

### 2.9 Optional encryption without ambiguous coverage

Users may save a project as a standard package or as an encrypted package. Encryption must protect the canonical project JSON, the portable Template Pack snapshot, and every packaged Resource and derivative. It must never imply that an encrypted project also protects the external originals from which Resources were imported or any exported PPTX, PDF, SVG, or PNG files.

### 2.10 No silent external file dependencies

Adding a file to a project copies its bytes into the project package. Slides, citations, extraction records, and MCP responses resolve stable Resource IDs inside the package, not filesystem paths. Original paths may be recorded as provenance but are never required to reopen, render, export, or share the project. Any future opt-in external-link feature must label the project non-portable and is outside the initial release.

### 2.11 Minimum necessary change

Cleanup changes the smallest property and scope that resolves a verified issue. A slide, table, figure, or object that already conforms remains untouched. A local difference that could represent purposeful technical emphasis, a prior editor decision, customer intent, or sponsor-template behavior is not auto-corrected; it is preserved or staged for review with an explanation.

### 2.12 Correct template before ORNL branding

Every imported deck is classified as current ORNL, older/modified ORNL, sponsor, custom, mixed, or unknown before cleanup. ORNL rules apply only to a deck whose selected target is an authorized ORNL Template Pack. A sponsor deck uses its supplied sponsor Template Pack and rule profile. Cross-template conversion is a separate explicit reflow request, never an inferred cleanup action.

---

## 3. Release scope

### 3.1 Required product capabilities

- Electron desktop application for macOS and Windows.
- React and TypeScript interface with a Vite-based local build.
- Local project library and versioned, self-contained `.pstudio` project-package format.
- Optional encrypted `.pstudio-secure` project package with encrypted autosave/recovery behavior.
- A Resources workspace for adding, previewing, organizing, tracing, and authorizing documents, data, images, audio, video, and imported presentation content.
- Autosave, atomic file replacement, recovery state, undo, and redo.
- Versioned local Template Packs.
- First Template Pack for the official ORNL 16:9 presentation system.
- Importable, project-local sponsor/custom Template Packs with explicit authority and support status.
- Multi-deck Batch Review workspace with per-deck target-template classification and isolated outputs.
- Read-only audit reports for template, font, table, figure, layout, and production consistency.
- Conservative Cleanup mode with confidence-scored, reviewable rule-fix batches.
- Approved style exemplars for tables, figures, and other bounded object families.
- Manual slide composition and layout selection.
- Document and manuscript intake with local extraction.
- PowerPoint intake and template-aware reflow.
- Audit, Conservative Cleanup, Reflow, Compose, Hybrid, and Manual work modes.
- Before/After proposal review and stale-write protection.
- Editable PowerPoint export.
- Multi-page PDF export.
- Per-slide and whole-deck SVG and PNG export.
- Local, model-independent MCP server.
- Source-based macOS and Windows setup scripts.
- Start scripts/launchers for macOS and Windows.
- A committed, verified installation-guide PDF.
- Automated schema, renderer, exporter, MCP, Electron, setup, and content-fidelity tests.

### 3.2 Explicit non-goals for the initial releases

- Replacing the full PowerPoint authoring interface.
- Heavily redesigning an imported deck during Conservative Cleanup.
- Rewriting, proofreading, or correcting visible content unless the user selects a content-authorized mode.
- Automatically applying ORNL branding to a sponsor or unknown template.
- Treating an isolated formatting difference as an error without considering prior intentional edits or semantic styling.
- Claiming that automated “technical checks” prove scientific, engineering, regulatory, sponsor, or editorial correctness.
- Perfect round-trip support for every PowerPoint object or feature.
- Editing PowerPoint animation timelines.
- Reconstructing SmartArt as fully editable Presentation Studio objects.
- Editing embedded OLE files, macros, 3D models, or arbitrary add-in content.
- Publishing or distributing decks automatically.
- Sending email, uploading files, or writing to collaboration systems.
- A hosted MCP endpoint or hosted presentation service.
- A built-in remote AI provider or provider-specific API key workflow.
- External linked-resource mode, automatic cloud-resource synchronization, or background URL downloading.
- Password-protecting or encrypting exported PPTX, PDF, SVG, or PNG files in the initial releases.
- DMG, PKG, MSI, EXE, Squirrel, Microsoft Store, or Mac App Store installers.
- Disabling Gatekeeper, SmartScreen, PowerShell policy, or other operating-system security controls.
- Claiming official ORNL approval, accessibility certification, or publication approval.

### 3.3 Repository data boundary

The repository is public at the time of this specification. The official ORNL `.potx`/`.pptx` template, extracted template artwork, and compiled production ORNL Template Pack must **not** be committed. An authorized user or administrator installs the template from an approved local file or managed local folder. The application stores the compiled pack in local application data and may snapshot it into a self-contained project when rights allow. Git contains only the pack schema, compiler and validation logic, non-sensitive metadata needed by that logic, and synthetic/public test fixtures. The following items must not be committed unless an authorized owner explicitly clears them:

- standalone protected ORNL logo artwork not already embedded in the authorized template;
- official ORNL presentation templates, extracted assets, or production Template Pack bytes;
- restricted or licensed fonts;
- non-public brand-source files;
- imported papers, manuscripts, source documents, or presentations;
- user project packages, JSON snapshots, local recovery files, previews, or exports;
- extracted source text, images, speaker notes, or document metadata;
- MCP connection files, tokens, receipts, or credentials.

Development fixtures must be synthetic, public, or explicitly cleared. The first implementation must add ignore rules and tracked-file/history scanning before real source intake is enabled.

---

## 4. Users and primary jobs

### 4.1 Presentation author

Creates a new deck from approved source material, selects the audience and objective, reviews the proposed narrative, and refines individual slides.

### 4.2 Presentation designer

Reviews hierarchy, imagery, typography, layout choice, crop quality, consistency, and template fidelity. Can override layout assignments and refine the design without editing the source document.

### 4.3 Technical editor or production reviewer

Imports one or more decks, selects the correct rule profile, identifies an approved table/figure exemplar, reviews font and consistency findings, protects already-cleaned slides, and approves only the cleanup changes that match the assignment.

### 4.4 Batch coordinator

Creates a review batch, imports multiple decks, classifies ORNL versus sponsor targets, monitors progress and exceptions, and exports separate cleaned copies plus audit reports without merging or overwriting sources.

### 4.5 Existing-deck owner

Imports a large deck, chooses a preservation policy, reviews batch reflow proposals, and resolves slides that cannot fit without an explicit content decision.

### 4.6 Content or domain reviewer

Checks assertions, evidence, numbers, quotations, terminology, source links, and locked-content diffs before a deck is exported or distributed.

### 4.7 AI-assisted user

Uses an MCP-capable desktop client to read an authorized project, inspect bounded source excerpts, and stage presentation changes for review inside the application.

---

## 5. Core workflows

### 5.1 Audit and conservative cleanup

1. Create a single-deck review job or add the deck to a review batch.
2. Import the PowerPoint as a read-only `import-origin` Resource; never alter the source file.
3. Classify the source as current ORNL, older/modified ORNL, sponsor, custom, mixed, or unknown and have a person confirm the target Template Pack.
4. Select a versioned Cleanup Rule Profile and, when supplied, designate approved table or figure exemplars.
5. Optionally add an earlier cleaned/approved deck as a `prior-approved-revision` Resource for three-way style comparison.
6. Inventory effective fonts, masters/layouts, local overrides, placeholders, tables, figures, charts, notes, comments, and render/layout findings.
7. Show a read-only audit before proposing changes, grouped by rule, slide, severity, confidence, and likely review effort.
8. Protect approved-as-is decks, slides, object families, or individual objects.
9. Stage bounded proposals such as a font mapping, table-style normalization, figure-treatment normalization, alignment repair, or placeholder correction.
10. Re-render every affected slide and reject any proposal that changes protected content, creates overflow/overlap, loses an object, or violates the selected template.
11. Review changes individually or as a rule-scoped batch; Apply, Reject, Exclude, or Defer remain human actions.
12. Export a new editable PPTX and a cleanup report; never overwrite the source.

Conservative Cleanup does not rewrite text, correct scientific content, change slide order/count, replace imagery, add icons, create assertions, or broadly remap layouts. Those actions require Reflow, Hybrid, or Compose scope. The default is no change when the application cannot distinguish an error from intentional formatting.

### 5.2 Multi-deck batch review

A review batch coordinates many independent deck jobs without merging their content or project histories:

1. Add multiple PPTX files or select a folder after a visible file-count and destination review.
2. Create one immutable source record and one isolated working project per deck.
3. Preflight all decks and show slide count, template classification, fonts, object support, warnings, render availability, and estimated exception volume.
4. Require target-template confirmation for sponsor, mixed, or unknown decks before any cleanup proposal.
5. Apply a batch rule profile only to compatible confirmed decks; deck-level overrides remain visible.
6. Queue parse, render, proposal, validation, and export work with pause, cancel, resume, per-deck checkpoints, and failure isolation.
7. Review high-confidence changes by rule across slides while retaining Before/After access to every affected slide.
8. Export separate named copies and per-deck reports to a user-selected folder after a collision/overwrite check.

The Batch workspace reports `Not scanned`, `Audited`, `Needs template decision`, `Ready for cleanup`, `Proposal ready`, `Needs manual review`, `Approved`, `Exported`, or `Failed`. A failed deck or slide cannot block the rest of the batch, and partial completion is never reported as batch success.

### 5.3 Compose from source

1. Create a presentation project.
2. Select the approved Template Pack.
3. Add one or more approved files to Resources; Presentation Studio copies them into the project package.
4. Review the local extraction report before enabling AI access.
5. Define audience, presentation job, desired outcome, length, and required content.
6. Allow an AI client to read bounded source excerpts for the current session.
7. Stage a deck outline containing the central takeaway, supporting messages, narrative order, and source references.
8. Review and approve the outline.
9. Stage slides using full-sentence assertion headlines, evidence, visuals, and source notes.
10. Review slides individually or in bounded batches.
11. Apply accepted proposals to the working project.
12. Resolve validation warnings and export.

Composition is generative. The UI must state clearly that source content may be summarized, reorganized, or rewritten within the authorized policy. Generated assertions must retain source references and must not invent claims, methods, results, qualifiers, or implications.

### 5.4 Reflow an existing deck

1. Import a PowerPoint deck as a read-only `import-origin` Resource whose bytes are copied into the project package.
2. Inventory slides, layouts, masters, text, notes, media, tables, charts, groups, and unsupported objects.
3. Choose a content policy and a slide-count policy.
4. Generate an import report with object-level support status and risks.
5. Select the target Template Pack.
6. Stage reflow proposals in resumable slide batches.
7. For each slide, classify content roles and select a compatible target layout.
8. Rebind content to template slots, normalize hierarchy, and adjust geometry.
9. Show Before and After alongside exact-content and object-preservation checks.
10. Apply, reject, remap, or defer each slide.
11. Export a new deck; never overwrite the imported source deck.

Strict reflow may change line wrapping, text-box geometry, crop windows, scale, alignment, spacing, type style, slide layout, and slide breaks. It may not add, delete, correct, merge, paraphrase, or reorder locked semantic content.

### 5.5 Hybrid revision

Hybrid mode applies policies at deck, slide, group, and element levels. Typical examples include:

- preserve scientific terminology and numeric results while allowing assertion headlines to be rewritten;
- preserve quotations and citations while condensing explanatory body text;
- preserve tables, charts, and conclusion language while allowing slide splits;
- preserve all visible text but permit alternative image crops and layout assignments.

Every proposed change must identify the policy that authorized it.

### 5.6 Manual design

Manual work must not require MCP or an AI model. The Slides workspace is a direct-manipulation editor rather than a preview-only gallery. Users can create slides, choose layouts, fill semantic slots, add supported native elements, edit text in place, move and resize objects, adjust crops, edit tables and notes, validate, save, and export.

The initial interaction set includes selection and multi-selection, drag, resize, rotate where the export representation supports it, keyboard nudge, duplicate, group/ungroup, lock, layer ordering, align, distribute, snap-to-grid, smart guides, rulers, safe areas, crop/fill/fit, style inspection, and accessible undo/redo. The inspector exposes only properties that have a defined canonical-scene and PowerPoint representation. Browser-only CSS effects cannot masquerade as editable slide features.

Every human action is a small validated command against an exact scene revision. Commands record affected stable IDs, Before/After values, policy checks, author, and timestamp. A command that changes locked content, damages a protected object, violates a Template Pack constraint, or creates an unresolved fit failure is rejected or surfaced as a deliberate policy decision. Consecutive pointer moves may coalesce into one undo step, but the final geometry and its validation remain inspectable.

Human and AI editing is cooperative rather than simultaneous multiwriter editing. While a person is actively manipulating an object, the app places a short-lived soft lock on it. An AI proposal always targets a known base revision; any intervening human or accepted AI edit makes the proposal stale and requires rebase or regeneration. The UI may show which slide or objects the AI is inspecting, but the AI cannot move the live selection underneath the user.

### 5.7 Anchored design feedback

A user can switch from **Edit** to **Comment** and click an object, select a text range, select a table cell/row/column, or drag a rectangular region. The resulting design thread is saved in the project and appears as a numbered pin on Source, Current, and Proposal views. Threads support plain-language instructions, optional drawing markup or reference attachment, open/resolved/reopened status, assignee, and reply history.

An anchor never relies on pixels alone. It stores the slide ID and revision plus the strongest available semantic target:

- stable element ID for an object;
- stable table ID plus row, column, and cell IDs for table feedback;
- text-object ID plus character range and selected-text hash for copy-level feedback;
- normalized slide-relative region and leader endpoint for a purely visual area;
- the original normalized bounding box, render hash, and a small project-local reference crop as a recovery fallback.

When layout changes move the target, the pin follows its stable target. When an object is replaced, a deterministic lineage map transfers the anchor only when the relationship is unambiguous. Otherwise the thread becomes `needs-reanchor` while retaining its original revision and crop; the app never silently attaches it to a nearby object.

Submitting a thread creates a scoped AI work request. The AI reads the current slide revision, the exact anchor, the thread, relevant structured geometry, and the revision-bound render. It then stages a bounded resolution proposal and replies with what changed. The UI shows Source/Current/Proposal, highlights the affected objects, and offers Apply, Reject, Revise, Resolve, or Reopen. A thread cannot become `resolved` merely because the AI responded; resolution requires an applied change or an explicit human disposition. Routine fixes requested by a thread should be handled without follow-up questions unless the content, template, technical meaning, or unsupported-object boundary is genuinely ambiguous.

### 5.8 First-class table design

Tables are structured slide elements, not generic groups or flattened pictures. Import retains native cell text/runs, row and column order, merged-cell topology, row heights, column widths, cell margins, alignment, fills, borders, semantic color roles, notes/provenance, and the original PowerPoint object identity when available.

The table editor supports object-, row-, column-, cell-, and range-level selection; in-place text editing; row/column resizing and balanced distribution; cell padding; horizontal and vertical alignment; header, stub, body, subtotal, total, and note roles; border/fill/type tokens; merge/unmerge when content policy allows; approved style presets; and **Match approved exemplar**, **Fit to width**, and **Balance columns** actions. Changes remain native and editable in the exported PowerPoint whenever the imported object is supported.

Table layout uses a constraint solver rather than one global font reduction. It measures every cell with the intended font and final width, respects merged-cell constraints, protects minimum readable type and padding, computes row heights from wrapped content, keeps header/body hierarchy visible, and checks the table against slide safe areas and adjacent objects. When a table cannot remain legible on one slide, the application proposes a compatible larger table layout or a continuation-slide split with repeated header semantics. It does not delete cells, hide rows, rasterize the table, or silently shrink below the profile minimum.

Table visual QA is stricter than ordinary shape QA. Source, proposal, and independent export renders are checked for clipped or missing cell text, unexpected wrap changes, unequal or inadequate padding, inconsistent fonts and alignments, broken merged cells, border discontinuities, poor contrast, semantically destructive color changes, unbalanced widths/heights, overly dense presentation, and separation from captions/sources. Exact cell-content and structure hashes must pass before Apply and again after PPTX export. An approved exemplar supplies bounded style roles, not arbitrary geometry; incompatible tables are adapted through the role system or become visible exceptions.

### 5.9 Manage project Resources

1. Add files by file picker, drag-and-drop, paste, or extraction from an imported PowerPoint deck.
2. Validate file type, actual media signature, size, hash, and safe archive structure locally before acceptance.
3. Copy accepted bytes into the project package and assign a stable Resource ID; never depend on the original path afterward.
4. Create local, versioned derivatives such as extracted text, document locators, table snapshots, image thumbnails, audio transcripts, or video poster frames when supported.
5. Assign one or more roles: grounding source, slide media, chart/table data, imported-deck origin, generated output, or reference-only.
6. Review provenance, rights/authorization acknowledgement, extraction status, where-used references, and portability status.
7. Choose which Resources and representation types may be available to MCP for the current application session.

Removing a Resource never deletes or changes its external original. A Resource used by a slide, citation, import record, or pending proposal cannot be removed until the user resolves those references. Identical files are deduplicated by SHA-256 while retaining distinct provenance records when needed.

---

## 6. Content policies

### 6.1 Operation scope and content policy

Each review job has one operation scope:

| Scope | Meaning |
| --- | --- |
| `audit-only` | Inspect and report; no cleanup proposal can mutate the working project. |
| `cleanup-only` | Preserve content, slide count/order, object identity, and intentional template structure; stage only properties authorized by the Cleanup Rule Profile. |
| `reflow` | Preserve protected semantic content while allowing authorized layout reassignment, geometry changes, and slide splitting. |
| `hybrid` | Apply different cleanup, reflow, editable, or generative permissions to explicitly selected fields or objects. |
| `compose` | Create new source-grounded presentation content under the generative policy and review gates. |

Each deck and element also supports one of these content policies:

| Policy | Meaning |
| --- | --- |
| `locked` | Preserve semantic content and object identity. Geometry and approved style changes may be allowed. |
| `reflow-only` | Preserve text/data exactly; allow layout reassignment, line rewrap, crop, style normalization, and authorized slide splitting. |
| `editable` | A person may edit content directly. AI changes still require proposal review. |
| `generative` | AI may summarize, restructure, or rewrite within the source-grounding and closed-content rules. |
| `preserve-object` | Keep the original PowerPoint object intact because it is unsupported or must remain unchanged. |

Deck-level defaults may be overridden only by a more restrictive policy unless a person explicitly changes the scope.

`cleanup-only` is the default operation scope for imported decks. Its permitted properties are an allowlist, not an invitation to make a slide more attractive. New visuals, content rewriting, slide splitting, broad layout replacement, and cross-template conversion are off.

When a person explicitly asks Presentation Studio to **improve the design, alignment, layout, or visual quality of every slide while preserving the wording**, that instruction selects the `reflow` scope with a `reflow-only` content policy after one target-template confirmation. This is the **Designer Cleanup** workflow. It does not require per-slide approval for ordinary geometry, alignment, typography, spacing, table/figure treatment, or selection of a compatible layout from the confirmed Template Pack. The application stages the complete visual proposal for review before Apply/export and asks only when exact content cannot fit safely, the target template is ambiguous, a technical or semantic choice is unclear, or an unsupported object requires a human decision.

Designer Cleanup is deck-wide. An AI client must not stop after normalizing fonts or changing only the easiest slides. Every in-scope slide receives an explicit visual inspection and either a design improvement or a recorded `approved-as-is` judgment. Every text box is checked for bounds, overflow, wrapping, line breaks, insets, alignment, placement, font family, font size, weight, color, contrast, and line spacing. Tables, figures, captions, native diagrams, and any semantically useful icons are evaluated as part of the same composition. Completion requires an independently rendered exported artifact, a deck contact-sheet review, and repair of every known clipping, overflow, accidental overlap, off-slide object, unreadable text, or unbalanced placement issue.

### 6.2 Exact-content verification

For locked and reflow-only content, Presentation Studio must retain:

- the exact Unicode text sequence, excluding layout-driven soft wrapping;
- paragraph order and list order unless slide splitting is authorized;
- numbers, units, symbols, formula text, labels, and qualifiers;
- table row/column order and cell values;
- chart series names, category labels, data values, units, and source notes;
- notes text and citations;
- image/media asset bytes or exact source asset hash;
- object reading order unless an authorized layout change requires an explicit recorded change.

The project stores content hashes and a transformation ledger. A proposal that changes protected content cannot be applied.

### 6.3 Fit conflicts

The application must resolve fit problems in this order:

1. choose a more suitable approved layout;
2. use a permitted alternate placeholder or image crop;
3. split content into continuation slides when the selected policy allows it;
4. move approved detail to notes only when the selected policy allows it;
5. ask for an explicit content-policy decision.

It must not silently shrink body text below the template minimum, hide content outside the slide, clip text, remove text, or rasterize text to make it fit.

In `cleanup-only`, a font or formatting repair that creates a fit conflict is downgraded from group-applicable to `review-required`. Cleanup must not solve that new conflict by rewriting, splitting, shrinking below the selected rule profile, or changing layouts.

### 6.4 Cleanup Rule Profiles

A Cleanup Rule Profile is a versioned, inspectable ruleset pinned to a review job. It contains:

- target Template Pack ID/version/hash and authority note;
- allowed font families, theme-font expectations, explicit font mappings, and exemptions;
- title, subtitle, body, caption, label, equation, and table-cell type rules;
- permitted theme colors and semantic-color exceptions;
- table and figure style rules or approved exemplar IDs;
- slide size, master/layout, placeholder, footer, page-number, safe-area, and template-furniture checks;
- alignment, spacing, line, border, crop, image-resolution, and distortion tolerances;
- minimum readable type and contrast rules;
- production, editorial-consistency, and domain-review check configuration;
- properties that may be proposed in a group and properties that always require slide-level review;
- profile owner, source, date, version, and project-specific overrides.

The shipped ORNL cleanup profile uses the approved ORNL Template Pack and Aptos styles. Sponsor profiles are derived from the supplied sponsor template plus an authorized project rule record; ORNL colors, typography, assertions, or layout rules are not injected into them.

### 6.5 Finding confidence and review behavior

Every cleanup finding is assigned one confidence class:

| Class | Requirement |
| --- | --- |
| `deterministic-safe` | An explicit rule is violated, the intended target value is known, protected content is unchanged, and rerendering introduces no fit or visual regression. It may be grouped for review but is never silently applied. |
| `review-required` | The target is plausible but the difference may be intentional, semantically meaningful, inherited, or layout-affecting. |
| `manual-specialist` | The object is unsupported, the correction requires PowerPoint/domain expertise, or the requested technical judgment cannot be proven from authorized inputs. |

Findings retain rule ID/version, source value, proposed value, inheritance origin, affected objects/slides, confidence reason, Before/After renders, validation results, and disposition. Rejected/excluded findings are remembered for the current project so they are not repeatedly proposed.

### 6.6 Font consistency

Font auditing resolves the effective typeface for every text run and table cell through theme, master, layout, placeholder, object, paragraph, and run-level inheritance. It distinguishes declared fonts from the font actually available/rendered on the current system.

For an ORNL cleanup job:

- Aptos and the supplied template styles are the target for ordinary presentation content;
- Century Gothic, Arial, and other legacy local overrides are reported with their source scope and proposed mapping;
- Cambria Math, symbol fonts, equation objects, icon fonts, approved sponsor marks, and other declared exemptions are never blanket-replaced;
- a replacement is group-applicable only when text, paragraph order, size policy, and rendering remain valid;
- every replacement is rerendered because changed font metrics may alter wrapping, height, or pagination;
- the exported PPTX is re-inspected for unresolved fonts and PowerPoint substitution.

The font report groups findings by family, effective/inherited source, deck/slide/object, text role, and fit risk so a reviewer can approve one safe mapping without accepting unrelated typography changes.

### 6.7 Approved style exemplars

A user may mark a table, figure, chart, or bounded object family from an imported slide or separate reference deck as an **Approved Style Exemplar**. The project copies the exemplar and records its source hash, object locator, approval note, target object type, and allowed scope.

For tables, the style fingerprint may include header/body typography, fills, borders, line weights, cell margins, banding, alignment, row heights, and semantic category treatments. Applying it must preserve every source cell value, merged-cell structure, row/column order, notes, and any color that carries meaning unless the reviewer explicitly maps that semantic role.

For figures, the fingerprint may include crop/fit behavior, border/line treatment, caption/source style, label placement, alignment, spacing, and resolution expectations. It never replaces figure content or converts a native chart to an image merely for consistency.

Compatibility is checked before a style proposal. Structurally incompatible tables/figures become exceptions rather than being forced into the exemplar.

### 6.8 Prior approved revision comparison

When an earlier cleaned deck is available, it may be added as a read-only `prior-approved-revision` Resource. Presentation Studio aligns current and prior slides/objects using persistent PowerPoint IDs when reliable, then content, layout, geometry, and media fingerprints with confidence reporting.

The comparison has three inputs: current customer content, prior approved formatting, and the current confirmed Template Pack/Rule Profile. A proposal may recover a specifically approved formatting property from the prior revision while preserving the current deck's exact text, data, notes, media, slide order, and customer-added objects. It must not replace a current slide wholesale or roll content back to the earlier file.

Unmatched, reordered, deleted, or newly added slides/objects are shown explicitly. Low-confidence matches cannot participate in grouped proposals. The prior file is evidence, not absolute authority; current rules and human decisions may supersede it.

### 6.9 Technical-check boundary

The client-facing phrase `technical checks` is divided into three visible categories:

1. **Production checks:** deterministic template, font, overflow, overlap, off-slide, placeholder, figure resolution/distortion, table/figure treatment, line/border, contrast, and export integrity checks. These can block export.
2. **Editorial consistency checks:** possible inconsistencies in acronym expansion, capitalization, terminology, units, symbols, table/figure numbering, captions, and references. These are findings only unless an editor explicitly authorizes a textual correction.
3. **Domain technical review:** scientific, engineering, regulatory, sponsor, method, equation, and result correctness. The AI may identify questions only when given authorized sources or a project-specific checklist; a qualified human decides them.

The UI and report never collapse these categories into one “technically correct” badge. A clean production audit is not evidence of domain correctness.

### 6.10 Initial cleanup finding catalog

The first ORNL Cleanup Rule Profile should implement at least these finding families; exact values come from the pinned Template Pack/profile rather than hard-coded visual guesses:

| Rule family | Example | Default behavior |
| --- | --- | --- |
| `template` | Wrong dimensions, master/layout mismatch, missing required furniture, or mixed template lineage. | Block cleanup until target is confirmed; propose no cross-template conversion. |
| `font-family` | Century Gothic, Arial, or another non-exempt effective font in ordinary ORNL content. | Propose the explicit Aptos mapping only if fit-safe; otherwise review. |
| `font-role` | Inconsistent family, size, weight, or spacing among equivalent titles, body copy, captions, or table cells. | Compare inheritance and repeated style clusters; review before normalization. |
| `table-style` | Compatible table differs from an approved header/body/border/alignment treatment. | Apply only approved exemplar properties; preserve data, structure, and semantics. |
| `figure-style` | Inconsistent crop, border, alignment, caption/source treatment, resolution, or aspect distortion. | Auto-detect; repair only bounded geometry/style properties with unchanged media. |
| `layout-fit` | Overflow, clipping, off-slide content, unintended overlap, or wrapped one-line title. | Production blocker; no hidden/shrunk-text workaround. |
| `placeholder` | Empty inherited placeholder or visible default prompt text. | Production blocker; fill/delete only through an authorized exact-object proposal. |
| `theme-style` | Nonsemantic local fill, line, border, or color override outside the confirmed profile. | Review; preserve possible category/technical meaning. |
| `editorial-consistency` | Possible acronym, capitalization, unit, symbol, caption, table/figure numbering, or cross-reference inconsistency. | Non-mutating editor finding unless text editing is separately authorized. |
| `domain-review` | Possible scientific, engineering, equation, method, result, regulatory, or sponsor issue. | Manual-specialist question only; no approval or automatic correction. |

---

## 7. Assertion-evidence composition model

Each authored slide has one primary narrative job and one primary claim. The structured content model should include:

- `assertion`: a complete, audience-facing sentence that states the slide's main idea;
- `supportingMessages`: two to four concise points when needed;
- `evidence`: source-grounded data, quotation, chart, table, example, or visual;
- `implication`: why the evidence matters to the intended audience;
- `visualIntent`: the role a visual must play rather than a request for decoration;
- `sourceRefs`: references to exact document excerpts, pages, figures, tables, or imported slide objects;
- `speakerNotes`: detail, methods, caveats, and `[Sources]` blocks;
- `contentPolicy`: the authorization applied to each field.

The assertion is not required in strict reflow when creating one would rewrite locked content. In that case, the original title remains and an optional assertion suggestion can be staged separately without altering the slide.

---

## 8. Template system

### 8.1 Template Pack

A Template Pack is a versioned local bundle containing:

- template ID and human-readable name;
- template kind: `ornl`, `sponsor`, or `custom`;
- source template filename and SHA-256;
- source/version date and authority note;
- slide dimensions and supported aspect ratios;
- master and layout relationship metadata;
- named layouts and preview thumbnails;
- placeholder IDs, types, geometry, z-order, and inheritance;
- semantic slot roles;
- typography and paragraph styles;
- theme colors and permitted style tokens;
- master/background assets and content-safe regions;
- layout capacities and compatibility rules;
- object and export support declarations;
- design rules and validation checks;
- logo state and artwork provenance;
- migration metadata for newer pack versions.

The original template remains read-only. A project records the exact Template Pack ID, version, and source hash used to create or reflow it.

### 8.2 ORNL baseline Template Pack

The initial authorized local reference inspected for this specification was `13_ORNL_Presentation_16x9_Template.potx` with SHA-256:

`8ab5ef02fb1ebf102790e762eab81e4feed01f4496775c3b854da81445252d62`

That inspected revision contains one slide master, 30 named layouts, and six guidance/example slides. The compiled Template Pack must inventory every layout in the active authorized revision; for this initial reference, that means all 30 layouts, including:

- five title layouts;
- one-column and key-image layouts;
- two-, three-, and four-column variants;
- green-bar variants;
- multi-image series layouts;
- 10- and 14-portrait series layouts;
- four conclusion layouts.

The Pack must preserve the official master and use Aptos with supplied styles. It must capture presentation guidance including full-sentence assertion headlines, concise supporting content, purposeful visuals, approved palette use, readable typography, and source/rights expectations.

The authorized template remains outside the public repository. It is installed from an approved local source and preserved byte-for-byte in protected local application storage. A local manifest records the template name, revision/version, effective date when known, source/owner note, import date, SHA-256, layout inventory, and compatibility status. The UI shows the active revision and warns when a project is pinned to a different revision.

Installing a newer authorized revision creates a new side-by-side Template Pack version; it does not overwrite the previous pack or silently migrate projects. Qualification must rebuild the complete master/layout catalog, render representative synthetic content through every layout, compare master/placeholder/theme fidelity, and verify editable PowerPoint open/render behavior. The known hash above is evidence for the inspected initial reference, not a permanent claim that it is the newest authorized ORNL revision.

### 8.3 Sponsor/custom Template Packs and classification

A sponsor or custom template may be compiled into a project-local Template Pack after the user confirms its authority and intended deck scope. It is copied into the self-contained project package when rights allow, remains read-only, and is not installed globally or committed to the repository unless separately authorized.

Template classification uses a structural fingerprint rather than filename or visible colors alone. Signals include slide dimensions, master/layout/theme hashes, placeholder and relationship structure, theme fonts/colors, known template media hashes, footer/page-number behavior, and source-package provenance. The result is one of `current-ornl`, `older-or-modified-ornl`, `sponsor`, `custom`, `mixed`, or `unknown`, with confidence and supporting evidence.

`older-or-modified-ornl`, `mixed`, and `unknown` always require a human target decision. Cleanup within an older approved template and migration to the current ORNL template are different operations. A project-local sponsor pack may use general production checks, but ORNL-specific typography, palette, logo, layout, or assertion rules remain disabled unless explicitly part of that sponsor's authorized rules.

### 8.4 Template precedence

1. A newer approved template supplied by the user or responsible owner.
2. The locally approved ORNL baseline Template Pack.
3. A project-pinned older Template Pack for faithful reopening and export.

Presentation Studio must not silently migrate a deck to a newer template version. Migration creates a reviewable proposal.

The selected source and target pack versions must remain visible in audit, reflow, review, and export. Sponsor/custom templates remain isolated from the ORNL layout catalog and are never converted to ORNL merely because an ORNL pack is installed.

### 8.5 Layout selection

The layout engine must score candidate layouts using:

- narrative role;
- number and type of content regions;
- headline length;
- body text volume;
- image orientation and crop tolerance;
- table/chart dimensions;
- required labels or captions;
- preserved-object constraints;
- template capacity limits;
- slide-count policy;
- recent slide silhouettes to avoid monotonous repetition.

The chosen layout, score factors, fit warnings, and rejected alternatives must be inspectable.

---

## 9. Project format

### 9.1 File identity

- Standard self-contained project suffix: `.pstudio`
- Encrypted self-contained project suffix: `.pstudio-secure`
- Standard MIME-style identifier: `application/vnd.presentation-studio.project+zip`
- Encrypted MIME-style identifier: `application/vnd.presentation-studio.project-encrypted`
- Top-level schema ID: `presentation-studio-project`
- Initial schema version: `1`
- Initial package-format version: `1`

The normal saved project is a package, not a loose JSON file and asset folder. A `.pstudio` file is a ZIP64-compatible container with a Presentation Studio manifest and restricted internal paths. Renaming it to `.zip` may support authorized troubleshooting, but normal users interact with it as one project file.

The canonical `project.json` remains available as an optional content-only diagnostic/interchange export named `.presentation.json`. That JSON snapshot is not the normal saved project, does not contain binary Resource bytes, and must be labeled **Not portable** unless distributed with a separately validated Resource bundle. MCP operates on the open project model and does not require this JSON export.

### 9.2 Logical package structure

```text
project.pstudio
├── manifest.json
├── project.json
├── templates/<pack-id>/
│   ├── pack.json
│   ├── layouts/
│   ├── theme/
│   └── media/
├── decks/<deck-id>/
│   ├── deck.json
│   ├── slides/
│   ├── findings/
│   └── proposals/
├── rules/
│   ├── profiles/<profile-id>.json
│   └── exemplars/<exemplar-id>.json
├── resources/
│   ├── blobs/<sha256>
│   ├── records/<resource-id>.json
│   └── derivatives/<resource-id>/<derivative-id>
└── provenance/
    └── transformations.jsonl
```

`manifest.json` identifies the package format, project schema, required members, byte lengths, media types, SHA-256 hashes, and feature flags. Internal member names use generated IDs or content hashes rather than user-supplied filenames. Entries are sorted and serialized deterministically where practical so package integrity, migration, and test results are reproducible.

The package includes portable snapshots of every exact compiled Template Pack needed by the project, including normalized master/layout/theme data and authorized template media. A source `.potx` may be included when its distribution rights allow it. Restricted fonts or other dependencies may be embedded only when their licenses permit it; otherwise the package must declare the dependency and approved fallback and may not claim a fully portable rendering state until it is resolved.

One `.pstudio` project may be `single-deck` or `review-batch`. A review-batch project embeds independent deck records, source Resources, rules, findings, proposals, and histories while deduplicating identical blobs. Each deck retains its own target Template Pack and exports separately. No member deck depends on the original source path or another external project file.

PPTX, PDF, SVG, and PNG exports are not copied back into the package automatically because they are reproducible derived artifacts. A user may deliberately add one as a Resource when it has an independent reference or provenance purpose.

### 9.3 Top-level project fields

```json
{
  "schema": "presentation-studio-project",
  "schemaVersion": 1,
  "project": {
    "id": "stable-project-id",
    "name": "Presentation title",
    "type": "single-deck | review-batch",
    "createdAt": "ISO-8601 timestamp",
    "updatedAt": "ISO-8601 timestamp"
  },
  "templates": [],
  "ruleProfiles": [],
  "styleExemplars": [],
  "settings": {},
  "resources": [],
  "decks": [],
  "designThreads": [],
  "transactions": [],
  "transformations": [],
  "validation": {}
}
```

Each deck record includes stable ID, source Resource ID, template classification/evidence, confirmed target Template Pack, operation scope, Cleanup Rule Profile, status, slide records, findings, proposal/review lineage, export history, and protected/excluded scopes. Batch-level defaults are copied into each member deck with visible overrides; changing a batch default never silently rewrites a reviewed deck.

### 9.4 Resource registry

Resources unify the former source and asset concepts. One Resource may have several roles without duplicating its bytes: `grounding-source`, `slide-media`, `chart-data`, `import-origin`, `prior-approved-revision`, `style-exemplar`, `template-source`, `generated-output`, or `reference-only`. Each Resource record includes:

- stable Resource ID and package member reference;
- display name, detected media type, byte size, and SHA-256;
- one or more Resource roles;
- import time and original filename;
- optional original-path provenance, disabled by default and never used for resolution;
- parent/child lineage for extracted figures, converted media, or generated derivatives;
- extraction status and extractor version;
- page, section, paragraph, figure, table, slide, sheet, timestamp, or frame locator index;
- supported AI representations such as metadata, text, table chunks, transcript, image preview, bounded audio preview, or video poster frames;
- slide, citation, import, and proposal `usedBy` references;
- content sensitivity/authorization acknowledgement;
- rights/source attribution and accessibility metadata when applicable;
- import warnings and unsupported-content inventory.

Accepted original bytes are copied into `resources/blobs/` before the add operation completes. Identical bytes share one blob by SHA-256, but may retain separate Resource records when their provenance or role differs. External originals remain untouched and may be moved, renamed, or deleted without breaking the project.

Derivatives required for useful portable behavior are also packaged and versioned. Examples include extracted document text and locators, spreadsheet value snapshots, image thumbnails, PDF page previews, transcripts, waveform metadata, and video poster frames. Regenerable renderer caches that are not required to reopen or understand the project are excluded from the durable package or marked discardable. A derivative always records the producing extractor/version and parent Resource hash.

Current-session MCP grants are runtime state, not saved Resource metadata. Reopening any project, including an encrypted project, returns all AI content permissions to off.

### 9.5 Slide record

Each slide includes:

- stable ID and order;
- target layout ID and source-layout mapping;
- narrative role;
- content policy;
- template slot bindings;
- freeform supported elements;
- speaker notes and `[Sources]` data;
- accessibility metadata and reading order;
- Resource and exact-locator references;
- original imported-slide/object references when applicable;
- content hashes;
- proposal/review lineage;
- open and resolved design-thread IDs;
- validation findings;
- rendered preview cache key.

Every media-bearing slide element refers to a stable Resource ID plus an optional derivative ID and crop/trim settings. It never stores a path to the original file. A proposal containing an unresolved Resource ID fails validation.

#### 9.5.1 Design-thread and edit-transaction records

A design thread includes stable ID, project/deck/slide IDs, source revision, current target revision, anchor type and payload, fallback normalized geometry, original render hash/reference crop, instruction and replies, author/assignee, open/resolved/reopened/needs-reanchor status, related proposal/transaction IDs, and timestamps. Text anchors include a bounded selected-text hash rather than duplicating unnecessary source text. Table anchors use stable table, row, column, and cell IDs rather than a human-facing address alone.

An edit transaction includes stable ID, base/result revisions, human or MCP-proposal origin, ordered commands, affected stable IDs, Before/After values, validation output, content-policy result, coalescing/undo metadata, and timestamps. Human direct edits and applied AI proposals both produce this record. Rejected proposals do not enter the working transaction history but retain their review lineage.

### 9.6 Supported element types

Initial schema element types:

- text and rich-text runs;
- image;
- SVG/vector image;
- shape;
- line and connector;
- table;
- native chart for approved chart families;
- group;
- template placeholder binding;
- slide background reference;
- preserved source object;
- unsupported-object placeholder with review status.

Native tables additionally carry stable row, column, and cell IDs; merged-cell spans; semantic roles; style-token references; exact-content and structure hashes; and fit/legibility findings. Table text remains addressable at the cell and rich-text-run levels for editing, annotation, validation, and PowerPoint export.

The Resource library may retain file types that are not yet placeable or exportable. Support states are explicit: `source-readable`, `previewable`, `placeable`, `pptx-editable`, `pptx-preserved`, `render-only`, or `unsupported`. Storing a Resource must not imply that every export format supports it.

### 9.7 Package integrity, portability, and safe opening

A valid package must:

- resolve every required project, slide, Template Pack, Resource, derivative, citation, and provenance reference inside the package;
- match every declared size and SHA-256;
- contain no symlinks, absolute paths, parent-directory traversal, device paths, duplicate ambiguous names, or members outside the approved structure;
- enforce limits for total expanded bytes, compression ratio, member count, member size, nesting, and extraction time before allocating or extracting;
- determine actual file type from signatures and safe parsing rather than trusting extensions alone;
- never execute macros, scripts, embedded executables, active document content, or media on import;
- use bounded streaming reads and private temporary storage with cleanup after failure;
- report unsupported, corrupt, truncated, or quarantined members without opening a partial working project as valid.

The application displays one of three portability states: **Self-contained**, **Self-contained with declared font fallback**, or **Non-portable**. Initial-release `.pstudio` saves must be in one of the two self-contained states. URLs and filesystem paths may appear as citations or provenance, but never as required rendering dependencies.

### 9.8 Geometry

The canonical scene uses slide-relative inches for interoperable authoring and export. Imported PowerPoint objects also retain their exact source EMU geometry for fidelity checks. Template slot bindings should be preferred over AI-authored freeform coordinates.

### 9.9 Optional encrypted project package

Presentation Studio supports two save modes:

| Mode | Contents | Normal use |
| --- | --- | --- |
| Standard `.pstudio` | Readable package containing canonical JSON, a portable Template Pack snapshot, Resources, required derivatives, and provenance. | Portable local work where operating-system and storage controls provide the required protection. |
| Encrypted `.pstudio-secure` | Authenticated encryption over the same logical package contents. | Portable or at-rest protection for the complete saved Presentation Studio project. |

A standard `.pstudio` is a readable ZIP-compatible package. Anyone who can access the file can potentially read its embedded documents, data, media, notes, and extracted derivatives. The first standard save, and every change from encrypted to standard mode, requires a clear readable-data warning and explicit acknowledgement.

The initial encrypted format should adapt the encryption profile already used by OrgChart Studio:

- AES-256-GCM authenticated encryption;
- a passphrase-derived key using PBKDF2-HMAC-SHA-256;
- at least 250,000 PBKDF2 iterations for format version 1;
- a fresh random 16-byte salt for every save;
- a fresh random 12-byte IV/nonce for every encrypted payload;
- format/version metadata bound as authenticated additional data;
- a minimum 12-character passphrase;
- authenticated failure for the wrong passphrase or modified ciphertext.

The encrypted format must be independently versioned so cipher, key-derivation parameters, and serialization can migrate without changing the presentation-project schema. Encryption parameters may be strengthened in a newer format version after compatibility and performance testing.

The outer envelope may expose only nonsensitive format/version and encryption parameters. Project name, slide titles, Resource filenames, notes, thumbnails, Resource bytes, Template Pack content, and manifest member names remain inside authenticated ciphertext.

Because presentation projects can be much larger than OrgChart backup fixtures, the implementation must not require a Base64 JSON copy of the complete encrypted payload in memory. Use a streaming or bounded binary container and atomic file replacement. The final format must define authenticated chunk/member handling, truncation detection, size limits, and cleanup after failed saves or opens.

The passphrase is not stored in the project, recovery file, logs, activity history, or MCP runtime data and is not recoverable by Presentation Studio. The derived key may remain in memory only while the encrypted project is open. Optional operating-system keychain support is a future, separately reviewed capability.

Opening an encrypted project does not enable MCP or Resource-sharing permission. Those controls remain off or scoped according to the normal current-session rules.

Saving an encrypted project must display these boundaries:

- Presentation Studio cannot recover a forgotten passphrase.
- The packaged copies of imported Resources are encrypted; the external originals from which they were imported are not.
- Exported PPTX, PDF, SVG, and PNG files are not encrypted by this option.
- Cloud-synced destinations will sync the encrypted package to that provider.

---

## 10. Resource ingestion and local derivation

### 10.1 Initial supported inputs

- PowerPoint `.pptx` and template `.potx`;
- Word `.docx`;
- PDF `.pdf` with embedded text;
- Markdown `.md`;
- plain text `.txt`;
- CSV `.csv` and Excel `.xlsx` for tables/chart data;
- PNG, JPEG, WebP, TIFF, and approved SVG images;
- WAV, MP3, M4A, and approved audio containers;
- MP4/MOV video using an explicitly supported codec profile.

The exact media/codec matrix must be pinned and tested before a file type is advertised. OCR for image-only PDFs and raster images, speech-to-text, and video transcription are not required for the first implementation unless vetted local extractors are selected. Reports must distinguish unsupported extraction from genuinely empty content.

### 10.2 Intake transaction

Adding a Resource is an atomic, cancelable transaction:

1. inspect the candidate without modifying it;
2. enforce file and archive safety limits;
3. stream the bytes into a temporary package member while hashing;
4. validate the detected format and create safe local derivatives;
5. commit the blob, Resource record, derivatives, and manifest update together;
6. leave the previous package valid and remove partial temporary content after cancellation or failure.

The UI shows the increase in packaged project size before accepting unusually large media. Initial releases always copy accepted Resources into the package; they do not offer a link-only shortcut.

### 10.3 Local extraction boundary

Extraction occurs locally. Before AI access is enabled, the app shows:

- Resource names, types, sizes, hashes, and roles;
- extracted page/section counts;
- supported and unsupported objects;
- whether images, media, notes, data, transcripts, or derivatives were retained;
- the exact permission scope available to MCP;
- a warning that returned text becomes part of the AI conversation.

Resource-content permission resets off at every application restart or project close.

### 10.4 Bounded Resource reads

MCP reads return an authorized, bounded representation with stable locators, never an external path and never arbitrary original file bytes by default. Text, table data, transcripts, image previews, audio previews, and poster frames have separate size, duration, or dimension limits. The server enforces per-call and per-session limits and records a local session-only access receipt.

### 10.5 Source-grounding requirements

- Generated assertions cite one or more exact source locators.
- Non-trivial claims and externally sourced assets receive `[Sources]` entries in speaker notes.
- Unsupported or ambiguous extraction receives `needs_review`; it is never promoted to confirmed evidence automatically.
- The application must not invent missing citations, results, methods, names, dates, or units.

---

## 11. Existing PowerPoint import, cleanup, and large-deck reflow

### 11.1 Import contract

PowerPoint intake is a layered preservation pipeline, not one parser call. A failure in comments, metadata, one object type, or one third-party library must not make an otherwise usable deck unreadable. Import uses these layers:

1. **Package preflight:** safely inspect the ZIP/OOXML package, content types, relationships, part sizes, compression ratios, signatures, and external links without interpreting every feature.
2. **Structural inventory:** map presentations, slides, masters, layouts, themes, notes, modern and legacy comments, people metadata, sections, media, fonts, and object-bearing parts.
3. **Native interpretation:** convert supported slide objects into the canonical scene while retaining each object's source part, relationship, stable import ID, and original XML locator.
4. **Preservation envelope:** retain source parts and relationships needed to round-trip objects that are not fully interpreted.
5. **Visual baseline:** render the source deck with Microsoft PowerPoint when available. If that is unavailable, use a qualified compatible renderer and label the baseline accordingly.
6. **Import report:** disclose object-level support, parser failures, renderer differences, substitutions, and decisions that require review.

The importer must not depend on a single high-level PowerPoint library. Package preflight and source preservation must still succeed when native interpretation fails on a valid but unfamiliar part.

### 11.2 Import inventory and support states

The importer must inventory every source slide and report:

- master and layout relationships;
- placeholders and inherited elements;
- text objects, rich-text formatting, language, and reading order;
- notes, modern threaded comments, legacy comments, and their anchors;
- images, SVGs, audio, video, and media relationships;
- tables, merged cells, charts, chart workbooks, and category styling;
- groups, connectors, equations, SmartArt, OLE, macros, 3D, animation, and other specialized content;
- font families, embedded fonts, substitutions, and missing-font risk;
- hidden slides, sections, object counts, and estimated processing cost;
- malformed, oversized, externally linked, active, or parser-rejected parts.

Each imported object receives one explicit state:

| State | Meaning |
| --- | --- |
| `editable-native` | Represented in the canonical scene and exported as an editable PowerPoint object. |
| `preserved-native` | Not fully editable in Presentation Studio, but original OOXML is retained for controlled passthrough. |
| `conversion-required` | A proposed edit requires a disclosed conversion and explicit approval because editability or native behavior may be lost. |
| `unsupported-blocking` | No safe automatic interpretation or preservation path has been proven; the native baseline remains visible and specialist review is required. |

Comment data is review metadata, not audience-visible slide content. Comment parts must not block slide import, must be preserved when technically possible, and are not disclosed through MCP unless the user grants that representation for the current session.

### 11.3 Conservative cleanup behavior

Conservative Cleanup is the default for an imported deck. It must:

- preserve the complete original deck as a read-only `import-origin` Resource;
- create a separate Presentation Studio project and export to a new file;
- preserve slide order, hidden state, notes, exact text, table-cell values, chart data, equations, media identities, and comments by default;
- compute pre/post hashes for protected text runs, normalized text, table cells, chart data, notes, and Resource bindings;
- retain the source master/layout/slide hierarchy and edit the inherited or local property at its correct scope rather than adding a visual overlay;
- identify whether a differing style value is inherited, a local override, part of a repeated style cluster, or a one-off outlier before proposing a correction;
- compare only against the confirmed deck-specific Template Pack, Cleanup Rule Profile, and approved exemplars;
- reuse authentic source images and supported chart/table data;
- promote referenced slide media to individually usable Resources only when requested; keep unused, master-only, and template-package media preserved but out of the normal Resource library by default;
- preserve semantic color encodings until a person confirms that a restyle will not change meaning;
- preserve unsupported objects intact when technically possible;
- flag every fallback before application and avoid silent rasterization;
- record every object-level transformation and its reason.

Cleanup does not select a new layout, split a slide, add a visual, change wording, or reshape the narrative. An assertion headline, shortened bullet, corrected label, new claim, or reordered fact is a content change, not cleanup. It remains unavailable until a person changes the operation scope.

### 11.4 Strict reflow behavior

When a person explicitly selects Reflow, the application may additionally:

- map objects to semantic roles without changing locked content;
- select compatible target layouts and allow a person to remap a slide;
- adjust text-box geometry, spacing, alignment, crop windows, and scale within policy;
- permit slide splitting only under the selected policy and without reordering semantic content;
- preserve unsupported objects intact or require one of the explicit review choices below;
- stage, but not apply, any content-authorized assertion or wording suggestion separately from the layout proposal.

Reflow remains exact-content by default. Changing the operation scope expands allowed layout transformations; it does not silently change the content policy.

### 11.5 Unsupported objects

For an unsupported object, the review offers only explicit choices:

1. preserve the original object in the exported PPTX when technically possible;
2. keep the source slide unchanged;
3. rasterize a visible representation, acknowledging lost editability;
4. replace it manually with a supported object;
5. defer the slide for external PowerPoint editing.

No choice is made silently. The exported-deck audit must identify every object that changed support state.

### 11.6 Large-deck processing

- Parse slides incrementally where library constraints allow.
- Build the outline and package inventory before full-resolution previews.
- Default AI/reflow proposal batches to 10–25 slides.
- Allow pause, cancel, retry, and resume.
- Isolate a failed slide or part instead of failing the whole deck.
- Persist batch checkpoints outside the project file until applied.
- Keep the UI responsive while parsing, rendering, and exporting.
- Report per-slide progress, parser, render, and policy state.
- Never treat partial batch completion as whole-deck success.

The first performance qualification should include a synthetic or cleared deck of at least 250 slides. Maximum practical file size and memory targets must be set from measured macOS and Windows benchmarks before a public support claim is made.

### 11.7 Import qualification fixtures

Private or uncleared customer decks are never committed as test fixtures. The repository must instead include synthetic decks that reproduce important structures without source content, including:

- a title plus editable 7-by-3 table with merged cells and semantic fills;
- an approved table exemplar plus compatible and incompatible target tables;
- an ORNL-template deck containing inherited Aptos, local Century Gothic/Arial overrides, equation/symbol exemptions, an already-cleaned slide, and a font replacement that would cause wrapping;
- a sponsor-template deck that proves ORNL cleanup rules cannot be selected or applied accidentally;
- a mixed 25-plus-slide deck containing native text, shapes, connectors, notes, images, tables, and image-heavy slides;
- modern threaded comments anchored to text, pictures, and table cells;
- a valid deck whose comment part triggers a controlled high-level-parser failure;
- unsupported and preserved-native objects needed to exercise every review choice;
- native PowerPoint and alternate-renderer baselines that intentionally expose a renderer difference.

The regression gate is not merely that a deck opens. It must prove package preservation, exact-content hashes, editable-object counts, comment survival, slide count/order, and exported-render comparison.

---

## 12. Application experience

### 12.1 Visual expression

Presentation Studio uses a balanced, application-focused ORNL expression:

- ORNL Green as the identity anchor;
- Hale Navy for workspace structure;
- Polar, Graphite, and Dark Matter for surfaces and text;
- restrained Energy or Forge for focus/review states when appropriate;
- square 90-degree corners on brand-created containers;
- Mulish for the application UI when available, with Aptos and Arial fallbacks;
- Aptos for presentation content and PowerPoint compatibility;
- Phosphor Regular icons when an icon materially clarifies an action;
- no logo artwork until approved production art is supplied and authorized for this use.

The interface should adapt the shared design language of OrgChart Studio and USA Map Studio: fixed desktop shell, strong product mark, square-edged workspaces, a focused central canvas, dedicated modes, inspector-based editing, clear status, and restrained MCP activity feedback.

### 12.2 Desktop shell

- Top bar: product identity, project name, save state, undo/redo, zoom, validation, and export.
- Start screen: **Clean existing PowerPoint** is the primary action; **Create from source materials** is a secondary action.
- Left workspace rail: Library, Batch, Decks, Slides, Rules, Review, Resources, Compose/Story, Export, Local AI control, and Settings.
- Center: slide canvas, outline, source comparison, or review workspace depending on mode.
- Right inspector: selected slide, slot, element, source, policy, layout, and accessibility controls.
- Bottom/status region: zoom, selected object details, processing status, warnings, and autosave state.

### 12.3 Batch workspace

The Batch workspace is the default production view for the initial client goal. It includes:

- Add Decks and Add Folder intake with a reviewed file list before import;
- one row per deck showing filename, slide count, detected/confirmed template, target rule profile, fonts, table/figure counts, support risks, status, and exception count;
- aggregate counts for slides audited, deterministic-safe findings, review-required findings, manual-specialist findings, approved changes, unresolved blockers, and exports;
- filters by template family, status, finding rule, severity, confidence, font, object type, and assigned reviewer;
- queue controls for Audit, Prepare proposals, Render, Validate, Pause, Cancel, Resume, and Retry failed items;
- a **Review by rule** view such as `Century Gothic → Aptos`, `Table exemplar A`, or `Low-resolution figure`, with all affected slide previews;
- a **Review by deck** view preserving normal slide order and source context;
- explicit `Approved as-is`, `Exclude from rule`, `Defer`, and reviewer-note controls;
- export planning with one destination, per-deck filenames, collision checks, completion state, and cleanup-report generation.

The batch never exposes one opaque `Fix all` action. A group proposal is defined by one rule version and one bounded property change. Slides that fail rerender or content-fidelity validation leave that group and become exceptions.

### 12.4 Resources workspace

Resources is a persistent left-side workspace inspired by a notebook-style source library, but designed for presentation production. It includes:

- Add Files, drag-and-drop, paste, and imported-deck extraction actions;
- filters for All, Documents, Images, Data, Audio, Video, Presentations, Prior Revisions, Style Exemplars, Generated, Used, Unused, and Needs Review;
- search across Resource names, locally extracted text, tags, captions, transcripts, table/sheet names, and provenance;
- compact rows/cards showing type, role, size, extraction state, usage count, warning state, and current-session AI-sharing state;
- a detail view with native-safe preview, metadata, provenance, rights/authorization acknowledgement, accessibility fields, derivatives, and exact `Used on` links;
- a document/data view with page, section, table, sheet, timestamp, or frame locators;
- drag-to-slide and Insert actions for placeable Resources;
- `Use as evidence`, `Use as visual`, `Use as data`, and `Reference only` role controls;
- Replace, Re-extract, Locate in deck, Duplicate record, and Remove actions with dependency checks;
- a visible packaged-size summary and portability status.

Adding a Resource copies it into the package. Replacing one creates a new content hash and a reviewable update to dependent slide references; it does not silently mutate every use. The original external file is never changed or deleted.

The Resource detail view includes a **Share with AI this session** control. A user selects exact Resources and allowed representations: metadata, extracted text, table chunks, transcript, image preview, bounded audio preview, or video poster frames. The app shows the selected count and estimated accessible content, provides a one-click revoke action, and resets all grants on restart or project close. Storage in the project, permission to use in manual slides, and permission to disclose through MCP are three separate states.

### 12.5 Slide workspace

- Slide thumbnail rail with sections, hidden-state indicator, warnings, and proposal state.
- 16:9 canvas with pan, zoom, fit, rulers, guides, safe margins, and template-slot visualization.
- Edit and Comment modes with visible keyboard shortcuts, undo/redo, transaction history, and an always-available path back to Source.
- Selection, multi-selection, move, resize, supported rotation, crop, align, distribute, group, lock, snapping, smart guides, keyboard nudge, and layer ordering.
- Direct in-place editing for supported text and native tables, with a property inspector that shows the eventual PowerPoint representation.
- Template-layout chooser with preview, capacity, and compatibility explanation.
- Resource/evidence inspector linking a slide back to excerpts, figures, tables, media, data, or imported objects.
- Source, Current, Proposal, and overlay comparison modes with synchronized zoom.
- Per-slide controls for goal, content policy, protected objects, allowed layout changes, and whether new visuals may be proposed.
- Effective-style inspector showing whether each font, fill, border, spacing, or layout value comes from theme, master, layout, placeholder, object, paragraph, or run level.
- `Approved as-is`, `Protect object`, `Use as style exemplar`, and `Exclude from current rule` controls.
- Clear distinction between inherited template objects, editable content, preserved source objects, and app-only editor overlays.
- AI activity indicators at deck, slide, and affected-object scope; a user selection remains stable while AI work is in progress.

#### 12.5.1 Anchored feedback and design threads

Comment mode provides point pins, object selection, text-range selection, table-cell/range selection, and region markup. Pins remain visible at useful zoom levels without entering exported slides. Selecting a pin opens its thread in the inspector and highlights the exact semantic target plus the original-revision fallback region.

The thread composer states what will happen: **Submit to AI** creates a scoped proposal request; **Save note** records feedback without invoking MCP. Open threads are filterable by slide, object type, assignee, status, and severity. Resolved pins collapse but remain recoverable, and comments are included in the project history and review report without being injected into source PowerPoint comments unless the user explicitly chooses that export.

#### 12.5.2 Table design mode

Selecting a table replaces the generic shape inspector with a table-focused surface:

- mini grid navigator and direct cell/range selection;
- row, column, merge, padding, alignment, and semantic-role controls;
- approved Template Pack table styles and approved project exemplars with a live proposal preview;
- content density, widest-cell, wrapping, minimum-type, contrast, and structural warnings;
- Fit to width, Balance columns, Distribute rows, Repeat header on continuation, and Restore source structure actions;
- Source/Current/Proposal cell-level diff plus the exact cell/merge preservation state.

The canvas offers optical guides for table edges, header baselines, caption/source separation, and neighboring-object alignment. Table QA findings link directly to the affected cell or edge, and an AI can receive a bounded table crop plus structured cells/geometry when that representation is authorized.

### 12.6 AI Collaborator and Canvas Observer

The AI Collaborator is an iterative design partner, not a replacement canvas. Its normal unit of work is a slide-scoped **work order** containing:

- the user's goal and selected content policy;
- source, current, and pending-proposal revision IDs;
- native source/current/proposal renders when available, with renderer provenance and clearly labeled deterministic fallbacks;
- stable object IDs, semantic roles, bounding boxes, text metrics, z-order, and template slots;
- authoritative layout findings and protected-content hashes;
- confirmed template classification, Cleanup Rule Profile, style exemplars, prior exclusions, and finding confidence;
- target-layout candidates and a bounded, authorized Resource shortlist;
- allowed operations, including whether slide splitting or new visuals may be proposed.

The normal collaboration loop is **Inspect → Diagnose → Propose → Materialize candidate PPTX → Native render → Compare → Revise → Review → Apply or Reject → Validate export**. After each proposal, the app automatically renders the candidate through PowerPoint when available, refreshes deterministic findings, and returns the visual comparison to the AI. A configurable iteration cap and unchanged-finding stop rule prevent an autonomous loop from running indefinitely.

When work begins from a design thread, the loop is **Anchor → Inspect current revision → Stage bounded fix → Render → Reply → Apply/Reject/Revise → Resolve/Reopen**. The app passes the thread's semantic anchor, original reference crop, and current mapped geometry together so the model does not have to infer what “this area” means. Before staging, the AI re-reads the current revision and must not edit an object with an active human soft lock.

Canvas observation combines pixels and structure:

- a PowerPoint-native full-slide image tied to an exact package/scene revision when available;
- a deterministic editor render tied to the same revision, labeled as an editable-scene observation rather than native truth;
- an optional 2× image, selected-object crop, and bounded tiles for dense slides;
- a deck contact sheet for narrative rhythm and consistency;
- geometry, overflow, overlap, off-canvas, minimum-type, contrast, image-resolution, crop, reading-order, and template-slot findings;
- source-versus-proposal and editor-versus-export visual comparisons.

Models that accept image content inspect the highest-authority render available and receive its renderer provenance. Models that cannot still receive the structured scene and findings. Model vision is advisory: deterministic layout checks and final exported-artifact inspection remain authoritative. A model must not call a design ready when it has inspected only an approximate renderer and the required native-render gate remains pending.

In `cleanup-only`, the AI may diagnose and stage only allowed rule changes. It cannot propose a new assertion, image, icon, layout, slide, or wording change. A user must visibly change the operation scope before generative tools become available.

### 12.7 Visual Needs and generated Resources

The application does not use the term `filler image` in its authoring model. An AI may create a **Visual Need** only when a visual has a stated communication job, such as explaining a process, locating a place, distinguishing categories, or supplying useful evidence. A text-only slide remains valid when no relevant, approved visual helps.

The Visual Need resolver uses this order:

1. reuse an approved project Resource;
2. resolve a semantic icon name from the vetted local Phosphor Regular catalog;
3. construct a supported native PowerPoint shape or diagram;
4. stage a generated-image or custom-vector brief;
5. leave the need unresolved rather than adding decoration.

A generated visual brief records:

- slide and object target, communication purpose, subject, and factual grounding;
- intended medium: icon, diagram, illustration, photo-like image, texture, or other;
- aspect ratio, target dimensions, crop-safe region, background, style, palette, and visual hierarchy;
- positive prompt, negative constraints, prohibited content, and text-rendering policy;
- alt text, rights/authorization state, sensitivity classification, and approval status;
- generator/provider, model/version when known, generation time, prompt, parent brief, and output hash after generation.

Prompts sent to an external generator must be content-minimized and independently authorized for that environment; selected slide text or Resources are not disclosed merely because a visual was requested. Generated output first enters a quarantined **Resource candidate** state. It is signature-checked, decoded with size limits, scanned/sanitized, previewed, and accepted or rejected by a person before it becomes a packaged Resource or is placed on a slide.

CSS may style the Electron interface but is not a slide interchange format. A CSS or web-font icon shown on the canvas must resolve to a specific packaged, sanitized SVG or supported native geometry before export. The AI requests an icon by semantic catalog ID rather than inventing arbitrary SVG. Custom SVG uses a strict safe subset with no scripts, animation, external references, embedded HTML, remote fonts, or active content; unsupported SVG is rejected or rasterized only after explicit review.

Image generation and binary handoff are capability-negotiated. Every MCP-capable model can stage a Visual Need and prompt; a client that cannot generate or return the approved bytes leaves a ready brief for manual generation/import. A supporting client may stage a bounded visual Resource candidate, but cannot silently add it to the library or the deck.

Visual Needs are disabled in `audit-only` and `cleanup-only`. They belong to Reflow, Hybrid, or Compose and never appear as an implied solution to a consistency finding.

### 12.8 Review workspace

The review workspace must show:

- source/original slide and proposed slide;
- content diff with locked-content status;
- object preservation status;
- layout change and reason;
- source grounding and unresolved claims;
- template, fit, contrast, crop, overflow, and accessibility findings;
- Before/After values for changed fields;
- Apply, Reject, Remap, or Defer actions;
- stale proposal state when the project changes.

Cleanup review supports two synchronized perspectives:

- **By rule:** review one bounded change across many slides, automatically separating render/content exceptions;
- **By slide:** review all findings and prior decisions in the slide's original deck context.

For cleanup proposals, Review also shows exact source-text, table-cell, notes, media, and comment preservation results. For new visuals, it shows the brief, generated asset, provenance, grounding, rights acknowledgement, alt text, and whether the visual adds an unsupported factual implication.

### 12.9 Accessibility

- Full keyboard access for core authoring and review tasks.
- Visible focus treatment.
- Accessible names and semantic controls.
- Logical reading order for slide objects.
- Alt-text fields for meaningful images and graphics.
- Color-independent warning and status semantics.
- Contrast checks against approved color combinations.
- No claim of formal accessibility certification.

---

## 13. Rendering and export

### 13.1 Hybrid scene, preservation envelope, and render observations

Template bindings and supported project elements resolve into one PowerPoint-aware `SlideScene`. The editor, validation engine, native PowerPoint adapter, and SVG/PNG/PDF exporters consume the supported scene properties. Imported objects that are not fully represented remain in a separate preservation envelope keyed to stable source IDs and OOXML relationships. Export adapters must merge scene edits with preserved native content without silently reinterpreting, dropping, flattening, or duplicating it.

The slide workspace uses a hybrid canvas:

- a native PowerPoint render is the preferred visual foundation;
- supported scene objects provide selectable and editable overlays;
- preserved-native objects remain visually faithful and visibly limited in edit scope;
- guides, findings, comments, and AI activity appear as editor-only overlays that never reach export.

Every render produces a revision-bound `RenderObservation` containing renderer family/version, authority level, source package hash, scene revision, output dimensions, raster/SVG hash, font/substitution report, object bounds, text metrics, and deterministic findings. Valid renderer families include at least `powerpoint-native`, `qualified-alternate`, and `studio-approximate`. This record lets the AI and human discuss the same canvas state and prevents a proposal from being evaluated against a stale or mislabeled screenshot.

### 13.1.1 Native render bridge

When Microsoft PowerPoint is available, Presentation Studio uses a local bridge or companion add-in to render exact Source, Current, Proposal, and Export revisions without exposing project content outside the workstation. Rendering occurs from an immutable source or temporary candidate copy; it never overwrites the imported original. The bridge returns bounded slide images plus renderer/version, slide dimensions, package hash, font substitution where available, and failure diagnostics.

If native rendering is unavailable, the UI and MCP expose the best qualified fallback with an explicit authority warning. A fallback may support continued editing, but final PowerPoint-fidelity status remains `Needs native review` until a PowerPoint render is inspected.

### 13.2 PowerPoint export

PowerPoint export must:

- preserve or clone the approved template master/layout hierarchy;
- use the correct 16:9 dimensions;
- use native text boxes and rich-text runs where supported;
- use native shapes, lines, tables, and approved chart families where supported;
- preserve editable images and SVG artwork as separate objects;
- retain meaningful object names for the Selection Pane;
- preserve speaker notes and `[Sources]` blocks;
- preserve hidden slide state;
- retain reading order where supported;
- preserve source objects selected for intact passthrough;
- record generation metadata without suggesting approval;
- export to a new file selected by the user.

The implementation may use PptxGenJS for native object generation, but template fidelity and imported-deck preservation require a separately tested OOXML/template adapter. Library choice must not weaken the master/layout contract.

When Microsoft PowerPoint is installed, each design iteration and the final local audit should render the candidate/exported PPTX through PowerPoint and compare it with the approved source/current/proposal observations. An alternate Office-compatible renderer is useful for portability testing but cannot by itself prove PowerPoint fidelity. Renderer disagreements are reported; they are not automatically attributed to source corruption.

### 13.3 PDF export

- Multi-page PDF with one slide per page.
- Match slide dimensions and visual content.
- Preserve selectable text when practical.
- Include no editor-only overlays.
- Verify every page render, page count, dimensions, fonts, and clipping.

### 13.4 SVG export

- Export current slide or all slides.
- All-slides export uses a chosen folder or ZIP bundle.
- Text remains text when fidelity and font availability permit.
- Embed or package required assets deterministically.
- Use stable element IDs and named groups.
- Exclude editor-only controls and guides.

### 13.5 PNG export

- Export current slide or all slides.
- Support at least 1×, 2×, and 4× dimensions.
- Use a deterministic sRGB pipeline.
- Preserve transparency only when the slide/background policy allows it.
- Inspect dimensions, alpha behavior, and representative pixel content.

### 13.6 Project package and JSON snapshot

- Save `.pstudio` and `.pstudio-secure` packages with atomic temporary-file replacement or another proven crash-safe transaction design.
- Validate the manifest, JSON schema, internal references, member sizes, hashes, and safe package structure before opening or replacing the working project.
- Preserve unknown forward-compatible fields where safe.
- Migrate earlier schema versions in memory and write the newer version only after user save.
- Maintain protected internal recovery state until a project path is selected.
- Offer `.presentation.json` only as an explicitly labeled content snapshot for inspection or controlled interchange; it is not a complete project backup and cannot substitute for the package.

### 13.7 Export encryption boundary

Normal PPTX, PDF, SVG, and PNG exports are readable derived artifacts even when their source project is encrypted. Before exporting from an encrypted project, the app must show a concise notice that the chosen output is not protected by the project passphrase. Password-protected Office/PDF export is outside the initial release and must not be implied by the encrypted-project setting.

### 13.8 Cleanup and batch reports

Each cleaned deck has an in-app audit record and may export a human-readable PDF plus machine-readable JSON report. The report includes source/output hashes and names, target Template Pack and Cleanup Rule Profile, slide count, rule-grouped applied changes, excluded/protected scopes, validation results, unresolved production blockers, editorial-consistency questions, domain-review questions, unsupported objects, renderer used, and final status.

The report does not claim technical or brand approval and does not reproduce full slide text, comments, or thumbnails by default. Optional visual evidence is selected explicitly because a report may be stored or shared separately from the project. A batch report summarizes per-deck outcomes without replacing the individual reports or treating partial export as success.

---

## 14. Electron and local architecture

### 14.1 Proposed stack

- Electron desktop shell.
- React and TypeScript renderer.
- Vite build.
- Zod or equivalent runtime schemas.
- Model Context Protocol TypeScript SDK.
- Node.js 22.13 or later as the source-install baseline.
- Dedicated worker threads/processes for parsing, thumbnails, slide rendering, and large exports.
- Renderer libraries selected per format and qualified against the shared scene.

Exact dependency versions must be pinned in the lockfile when implementation begins.

### 14.2 Security posture

- Renderer sandbox enabled.
- Context isolation enabled.
- No Node.js access from page content.
- Narrow preload API with explicit IPC channels.
- Navigation and window creation denied unless intentionally handled.
- Application network policy denies non-loopback requests by default.
- Local service binds only to `127.0.0.1` on an ephemeral port.
- Every launch creates a new random session token.
- Private connection file uses owner-only permissions where supported.
- Connection file removed during normal shutdown and treated as stale after the app exits.
- No direct database, project-package, packaged-Resource, or external-original manipulation by MCP clients.
- No passphrase, derived key, or decrypted project content in logs, crash metadata, analytics, MCP receipts, or connection files.

### 14.3 Data locations

Working transaction state, recovery state, discardable preview caches, and MCP runtime files live in the current user's Application Support/AppData location or a user-selected approved folder, never inside the Git repository by default. Durable Resource copies live in the user-selected `.pstudio` or `.pstudio-secure` package.

Project packages selected by the user may live elsewhere. The imported external originals remain wherever the user placed them, but Presentation Studio does not depend on them after intake and does not relocate, modify, back up, or delete them.

For an encrypted project, autosave and recovery payloads containing project content must remain encrypted. Decrypted project text, notes, thumbnails, Resources, and derivatives must not be written to persistent temporary files or an unencrypted preview cache. The implementation must describe any operating-system or third-party thumbnail behavior it cannot control rather than claiming complete device-level secrecy.

### 14.4 Autosave and recovery

- Working changes update internal recovery state atomically.
- A bound project package checkpoints validated changes after a coalesced autosave interval and always before clean close.
- Interrupted writes retain the previous valid file.
- Startup offers recovery when internal state is newer than the bound project.
- Undo/redo operates on validated project transactions.
- Applying an MCP proposal is one undoable transaction.
- Encrypted projects write encrypted recovery state and require the project passphrase before recovery content can be opened after restart.
- Changing from standard to encrypted save mode, changing a passphrase, or returning to standard mode is an explicit Save As transaction that rewrites the complete package atomically.

Large immutable Resource blobs must not be Base64-encoded into JSON or fully buffered in renderer memory. Package reads, writes, copies, hashing, encryption, and export use bounded streams in worker processes. Implementations may reuse unchanged package members while writing a new atomic checkpoint, but must never sacrifice manifest verification, deletion semantics, or crash recovery merely to avoid copying bytes. The UI distinguishes `Changes recovered`, `Saving metadata`, `Adding Resource`, `Checkpointing package`, and `Saved` states so a large media import is not represented as complete prematurely.

---

## 15. MCP server

### 15.1 Compatibility objective

The MCP server must be model- and vendor-independent. It exposes standard MCP tools with JSON Schema-compatible inputs/outputs and standard MCP Resources where the client supports them, with no dependency on provider-specific message formats. Any desktop AI client that supports local STDIO MCP tools and the required schemas may connect.

Compatibility does not guarantee that every client or model can consume the same document size, image type, audio content, or tool-result volume. The server advertises representation capabilities and uses bounded results, pagination, stable IDs, and explicit continuation tokens. Text and structured-data fallbacks remain available when a client cannot consume visual or audio content blocks.

At session start, the server records whether the client can receive image content, resource links, embedded Resources, audio, and experimental long-running task results. Presentation Studio does not require experimental MCP features for core editing. All visual-observation tools have a structured fallback, and all long operations remain resumable through ordinary status tools when richer client support is absent.

### 15.2 Runtime design

- Local STDIO MCP process launched by the AI client.
- MCP process reads the private app connection file.
- MCP connects to the open Electron app through loopback HTTP.
- Electron app must be open for project tools to work.
- MCP refuses non-loopback application URLs.
- App can pause MCP without removing client configuration.
- Resource-sharing permission is controlled per Resource and representation and resets each app session or project close.
- UI shows current MCP scope, authorized Resource count, representation types, active operation, and recent bounded receipts.

### 15.3 MCP Resource model

For clients that implement MCP Resources, the server exposes authorized project Resources through stable URIs such as `presentation-studio://project/<project-id>/resource/<resource-id>/<representation>`. Resource listing and reading return only the representations authorized in the open app session. The server never exposes a filesystem path or package-internal path.

The split is intentional: standard MCP Resources provide read-only, application-controlled context, while MCP tools support model-invoked search and proposal actions. Compatibility read tools mirror essential retrieval so a tools-capable client that does not surface MCP Resources can still work within the same permissions and bounds.

Representations include:

- `metadata`: display name, type, size, roles, hash, provenance summary, usage, and support state;
- `text`: bounded extracted text with stable page/section/paragraph/slide locators;
- `table`: schema and bounded rows/cells from approved data snapshots;
- `transcript`: bounded timestamped audio/video transcript segments;
- `image-preview`: a size-limited raster preview with dimensions and crop context;
- `audio-preview`: a duration- and byte-limited approved audio clip when the client supports audio content;
- `poster-frame`: one or more size-limited video frames with timestamps;
- `slide-render`: a revision-bound, size-limited source/current/proposal slide image plus render metadata;
- `deck-contact-sheet`: a bounded montage of slide renders for consistency and narrative review.

Full original Resource bytes, complete archives, and unbounded video are not exposed through MCP V1. An AI proposal can select and place an already packaged Resource by stable ID without receiving the original bytes. The application validates the ID, permitted use, media support, and proposal revision before review.

### 15.4 Read tools

Initial proposed tools:

| Tool | Purpose |
| --- | --- |
| `get_app_status` | Report open project, save/portability state, MCP scope, authorized Resource counts, and pending proposal. |
| `get_design_contract` | Return the mandatory model-independent Designer Cleanup, content-preservation, autonomy, ORNL-brand, and independent-render QA instructions. |
| `list_projects` | List local presentation projects without returning complete deck content. |
| `get_batch_summary` | Read member deck IDs, statuses, confirmed template targets, rule profiles, slide counts, and categorized finding counts for the open review batch. |
| `get_deck_audit` | Read one deck's template classification, object support, production findings, confidence, and unresolved decisions without changing it. |
| `get_slide_design_context` | Read exact protected slide text plus bounded typography, object counts, warnings, and findings for up to 10 consecutive slides after session authorization. |
| `get_cleanup_findings` | List bounded findings filtered by deck, slide, rule, severity, confidence, object type, or disposition. |
| `get_font_inventory` | Read effective/declared font families, inheritance origin, roles, exemptions, affected object IDs, and fit risk. |
| `get_style_exemplars` | Read approved exemplar metadata, compatible object families, permitted style properties, and source provenance. |
| `get_prior_revision_comparison` | Read bounded current/prior slide/object matches, confidence, formatting differences, and current-content protections. |
| `get_cleanup_rule_profile` | Read the exact pinned cleanup rules, target Template Pack, exceptions, and group-review permissions for one deck. |
| `get_project_outline` | Read slide IDs, order, titles, narrative roles, layouts, policies, and warning counts. |
| `get_slide` | Read one authorized slide with stable element IDs and source references. |
| `get_slide_geometry` | Read stable object IDs, semantic roles, bounds, text metrics, z-order, template slots, and deterministic findings for one revision. |
| `get_slide_render` | Return a revision-bound Source, Current, Proposal, or Export render plus renderer provenance/authority as image content/resource link when supported, with a structured fallback. |
| `get_deck_contact_sheet` | Return a bounded, revision-bound montage and slide index for deck-level rhythm and consistency review. |
| `compare_slide_renders` | Return source/current/proposal or editor/export comparison metrics, difference image when supported, and categorized findings. |
| `reject_design_proposal` | Reject only the AI's currently pending draft after an authoritative Current/Proposal comparison, recording a concrete visual-regression rationale and raster evidence without applying, saving, exporting, or changing source bytes. |
| `get_layout_candidates` | Return ranked compatible Template Pack layouts with semantic-slot mappings, capacity, rejected alternatives, and fit risks for one slide/work order. |
| `run_visual_preflight` | Run revision-bound geometry, content, template, native-render, and cross-render checks and return exact blocking findings without changing state. |
| `list_design_threads` | List bounded open/resolved design threads by deck, slide, object type, assignee, and status without returning unrelated slide content. |
| `get_design_thread` | Read one thread, its exact semantic anchor, original revision/crop metadata, current anchor mapping, replies, and related proposal/transaction lineage. |
| `get_table_design_context` | Read one authorized native table's stable row/column/cell IDs, merged topology, exact cell text hashes/authorized text, geometry, semantic roles, styles, density, and deterministic fit findings. |
| `get_template_catalog` | List installed Template Packs and layout capabilities. |
| `get_layout` | Read one layout's semantic slots, capacity, and constraints. |
| `list_visual_catalog` | Search approved local icons, native diagram primitives, and already packaged visual Resources by semantic purpose. |
| `list_resources` | List authorized Resource metadata, roles, hashes, usage, representation support, and permission state. |
| `search_resources` | Search authorized extracted text, data labels, transcripts, tags, and metadata with bounded results. |
| `get_resource_metadata` | Read one authorized Resource's metadata, provenance, usage, locators, and derivative inventory. |
| `read_resource_text` | Return bounded extracted text or transcript with stable locators for an allowed Resource/session. |
| `read_resource_data` | Return bounded sheet/table schema and rows/cells for an allowed Resource/session. |
| `get_resource_preview` | Return an authorized size-limited image/audio preview or video poster frame when the client supports that content type. |
| `get_import_report` | Read supported/unsupported object findings for an imported presentation Resource. |
| `validate_project` | Run authoritative package, policy, Resource, grounding, and layout validation without changing state. |

### 15.5 Proposal tools

Initial proposed tools:

| Tool | Purpose |
| --- | --- |
| `stage_deck_outline` | Stage a source-grounded narrative outline for review. |
| `stage_slide_create` | Stage a new slide using an installed layout and semantic bindings. |
| `stage_slide_update` | Stage changes to one existing slide using exact stable IDs. |
| `stage_slide_design` | Stage one coordinated slide-level design proposal across supported objects, semantic roles, and a selected approved layout while preserving the active content policy. |
| `stage_alignment_pass` | Stage bounded align/distribute/baseline/safe-area corrections across explicit stable object IDs on one slide. |
| `stage_text_fit_update` | Stage bounded text-frame geometry, inset, wrapping, line-spacing, and fit-safe type changes without changing protected wording. |
| `stage_figure_treatment` | Stage bounded image/caption/source crop, alignment, spacing, and approved frame treatment without changing media identity or protected labels. |
| `stage_design_thread_resolution` | Stage a bounded fix and reply for one open thread against its current mapped anchor and exact scene revision; it cannot mark the thread resolved or apply the change. |
| `stage_table_update` | Stage native table role, geometry, padding, alignment, and approved style changes against explicit table/row/column/cell IDs while preserving protected cell content and structure. |
| `stage_cleanup_rule_batch` | Stage one versioned cleanup rule and one bounded property change across explicit object IDs, separating rerender/content exceptions. |
| `stage_style_exemplar_apply` | Stage approved exemplar properties onto explicit compatible table/figure objects while preserving protected content and semantics. |
| `stage_review_finding` | Stage a non-mutating editorial/domain question with exact object/slide evidence and category; it cannot represent a technical approval. |
| `stage_slide_reflow` | Stage one imported slide against a target layout and content policy. |
| `stage_deck_reflow_batch` | Stage a bounded batch of slide reflows; never the entire massive deck in one opaque proposal. |
| `stage_layout_change` | Remap one slide while preserving its applicable policies. |
| `stage_visual_need` | Stage a purpose-driven visual brief, including grounding, prompt, constraints, target geometry, rights state, and alt text; it does not generate or place an asset. |
| `stage_catalog_icon_insert` | Stage insertion of one vetted semantic icon by catalog ID, resolving it to packaged safe SVG or supported native geometry. |
| `stage_visual_resource_candidate` | Capability-negotiated: stage a bounded generated image or safe-vector payload plus required provenance in quarantine; it is not a Resource until a person accepts it. |
| `stage_resource_binding_update` | Add or correct Resource/locator references without changing unsupported source facts or Resource bytes. |
| `replace_project_draft` | Stage a complete validated candidate JSON model using an exact stale-write guard and existing Resource IDs; it cannot add, omit, delete, or mutate packaged Resource bytes or the protected Resource registry. |

Every proposal tool requires `expectedUpdatedAt` or an equivalent project revision. The server rejects stale proposals.

Tool availability is operation-scope aware. In `audit-only`, all mutating proposal tools are unavailable. In `cleanup-only`, only `stage_cleanup_rule_batch`, bounded `stage_table_update`, `stage_style_exemplar_apply`, `stage_review_finding`, and other explicitly non-generative cleanup tools are advertised; generic slide replacement, reflow, layout, outline, visual, and Resource-candidate tools are unavailable. `stage_design_thread_resolution` inherits the underlying thread target's scope and exposes only operations allowed by that content policy. Server-side policy validation remains authoritative even if a client calls a cached or previously advertised schema.

### 15.6 MCP write boundary

MCP V1 does not:

- Apply proposals, accept proposals, or reject a person's/previously accepted proposal. The AI may reject only its own currently pending design draft after an authoritative native comparison so routine failed iterations do not become user approval work.
- Choose a save path.
- Export files.
- Import, replace, delete, or return original Resource bytes. A staged generated-visual candidate is a quarantined proposal, not an imported Resource, and cannot reference an arbitrary path or remote URL.
- Modify Template Packs.
- Confirm or change a deck's target template, operation scope, Cleanup Rule Profile, style-exemplar approval, or `Approved as-is` protection.
- Change storage settings.
- Enable Resource extraction or AI sharing.
- Publish, upload, email, or share a deck.

These remain visible user actions inside Presentation Studio.

`stage_visual_resource_candidate` is optional and strictly bounded. It accepts only allowlisted declared media types, rejects mismatched signatures and decompression hazards, sanitizes SVG, requires prompt/provenance/rights fields, and stores candidate bytes in protected transient proposal state. Accepting the candidate in the application performs the normal Resource intake transaction; rejecting it removes the transient candidate. Clients without safe binary handoff use `stage_visual_need` and the normal Add Resource workflow.

### 15.7 Activity and review

When an MCP tool is running, the application shows a restrained, color-independent activity state naming the operation. On completion, the pending proposal appears in Review. Apply/Reject decisions create a bounded local record containing project/slide IDs, affected fields, policy results, and timestamps, but not the user's prompt or raw source excerpts.

---

## 16. Source-based setup and launch

### 16.1 Distribution policy

Presentation Studio will use public source-based setup while signing and installer distribution are out of scope. It must not create or advertise DMG, PKG, MSI, EXE, Squirrel, Store, or similar installers.

### 16.2 Required setup files

Planned files:

```text
scripts/setup-macos.zsh
scripts/setup-windows.ps1
scripts/start-macos.zsh
scripts/start-windows.ps1
scripts/start-electron.mjs
scripts/configure-presentation-mcp.mjs
Start-Presentation-Studio.command
Start-Presentation-Studio.cmd
```

### 16.3 Setup behavior

Both platform setup scripts must:

- run as a normal user without administrator privileges;
- verify the repository origin and selected revision;
- safely fast-forward a clean `main` checkout when updating;
- refuse to overwrite a dirty checkout or unexpected branch;
- use an existing compatible Node runtime or install a private pinned runtime under `.runtime`;
- verify any downloaded runtime against an authoritative checksum;
- install exact lockfile dependencies;
- build the renderer and local services;
- run automated tests appropriate for installation;
- run a hidden Electron smoke test that exercises startup, storage, renderer, and local bridge;
- optionally configure the MCP server after an explicit prompt or documented environment switch;
- record the installed revision and checksums;
- leave user project packages, embedded Resources, external originals, and exports untouched during update/repair;
- provide an actionable failure message and nonzero exit status;
- never disable operating-system security controls.

### 16.4 MCP setup

Setup should offer MCP configuration as an optional step. It must identify the server as `presentation_studio`, use an absolute project/runtime path, and document how to remove the configuration. Skipping MCP must not prevent the application from running manually.

### 16.5 Update and repair

Rerunning the same setup command should update or repair a clean source installation. Update must preserve `.runtime` when compatible, local Application Support/AppData, `.pstudio`/`.pstudio-secure` packages and their embedded Resources, external originals, and user-selected output folders.

### 16.6 Uninstall

The installation guide must distinguish:

- removing MCP client configuration;
- removing the cloned source/runtime folder;
- optionally removing local Application Support/AppData;
- preserving or deleting user-selected project files and exports.

Data deletion must never be bundled into a normal update or repair.

---

## 17. Installation-guide PDF

### 17.1 Required artifact

The repository must include:

`docs/Presentation-Studio-Installation-Guide.pdf`

The committed PDF is a product deliverable, not a screenshot-only placeholder.

### 17.2 Required guide content

- Product and draft-status description.
- System requirements.
- macOS source setup with copy/paste commands.
- Windows source setup with copy/paste commands.
- First launch and later launch instructions.
- Optional MCP configuration and client restart.
- Confirmation that Presentation Studio must be open for MCP tools.
- Explanation of the self-contained `.pstudio` package, embedded Resources, packaged-size implications, portability status, and the limited diagnostic role of `.presentation.json`.
- Explanation of per-Resource MCP permissions and exactly which metadata, text, data, transcript, or preview can enter an AI conversation.
- Update and repair instructions.
- MCP removal and application uninstall instructions.
- Project-package, external-original, recovery, cache, and MCP runtime storage locations.
- Standard versus encrypted project save behavior, the readable embedded-Resource warning, passphrase loss, and the boundary around unencrypted exports and external originals.
- Troubleshooting for Node, Git, blocked scripts, port/runtime files, and smoke-test failures.
- Explicit statement that no signed installer is provided.
- Data-safety warning for imported source documents.

### 17.3 PDF production requirements

- Generate from a committed, editable documentation source.
- Use a deterministic build command.
- Use ORNL-approved colors and square-cornered containers.
- Use Aptos or a compatible documented fallback.
- Do not include protected logo art unless approved production artwork is supplied and cleared for repository distribution.
- Keep copy exact to the implemented scripts; regenerate after commands change.
- Render every page and inspect visually.
- Verify page count, media box, embedded/selectable text, command text, links, clipping, and font substitution.
- Test every documented setup command independently in a fresh environment before making a public installation claim.

---

## 18. Validation and quality gates

### 18.1 Project validation

- Package-format version and manifest schema.
- Schema and version.
- Stable unique IDs.
- Valid slide order and layout references.
- Valid element geometry and z-order.
- Required template bindings.
- Resource, derivative, Template Pack, and exact-locator references.
- Content-policy compatibility.
- Source-reference integrity.
- Missing or unfilled placeholders.
- Unsupported-object decisions.
- Notes and source provenance.
- Member sizes and SHA-256 hashes, safe internal paths, duplicate-name rejection, and expanded-size/compression limits.
- Self-contained portability status and declared font fallbacks/dependencies.
- Encrypted envelope version, declared parameters, authenticated manifest, and complete Resource/Template Pack recovery after decryption.

### 18.2 Content-fidelity validation

- Exact locked text comparison.
- Source-versus-output slide count/order, hidden state, notes, and comments comparison for Cleanup scope.
- Table and chart data comparison.
- Image/media SHA-256 comparison.
- Notes/citation comparison.
- Transformation authorization.
- `Approved as-is`, protected, and excluded object/slide hashes unchanged.
- No hidden clipped text or off-slide content.
- No unapproved move from visible content to notes.

### 18.3 Cleanup and batch validation

- Every deck has a human-confirmed target Template Pack before cleanup.
- ORNL rules are absent from sponsor/custom decks unless explicitly included in that deck's approved profile.
- Effective-font inventory covers themes, masters, layouts, placeholders, shapes, runs, and table cells; declared exemptions remain untouched.
- Each group proposal has one pinned rule version and one bounded property change.
- Applied font mappings leave no unauthorized family, substitution, overflow, clipping, or wrapping regression.
- Table/figure exemplar proposals preserve protected content, object type, merged structure, media identity, and semantic styling.
- Native-table validation checks every cell/range for clipping, missing text, unexpected wrapping, minimum type, padding, alignment, merge integrity, border continuity, contrast, semantic color preservation, and balanced row/column geometry.
- Table continuation proposals preserve exact row order and cell content, repeat the intended header semantics, identify continuation slides, and never hide or silently drop rows.
- Prior-revision proposals use only high-confidence matched objects, transfer only authorized formatting properties, and preserve all current content/customer-added objects.
- Already-cleaned/approved-as-is slides and objects remain byte/content/style equivalent at their protected scope.
- Each member deck retains isolated status, target, proposal history, output name, and failure state.
- A failed deck/slide is isolated; batch totals distinguish complete, partial, unresolved, and failed work.
- Production blockers, editorial-consistency questions, and domain-review questions remain separate in UI and reports.
- Original source files remain unchanged and every output uses a new path after an overwrite/collision check.

### 18.4 Layout and visual validation

- No unintended overlaps.
- No clipped or overflowing text.
- No one-line title wrapped unexpectedly.
- No body/caption text below template minimums.
- Valid image crops and sufficient resolution.
- Consistent alignment and spacing.
- Table headers, bodies, captions, sources, borders, cell padding, row heights, and column widths form a legible coherent system at delivery scale.
- Theme-color compliance.
- Contrast and color-independent meaning.
- Square-cornered brand-created containers.
- Required master/layout furniture present.
- No empty inherited placeholders in final PPTX XML.

### 18.5 Export validation

Each export format is verified independently:

- PPTX opens, preserves master/layouts, retains expected editability, and renders every slide.
- PDF page count, size, text, and rendering match the project.
- SVG parses, resolves assets, uses expected dimensions, and excludes editor overlays.
- PNG dimensions, color mode, alpha behavior, and non-empty pixels are correct.
- A standard package reopens without its external originals and produces the same validated project, Resource, derivative, and Template Pack hashes.
- A `.presentation.json` content snapshot matches the canonical model but is clearly identified as non-portable and incomplete without Resource bytes.
- Encrypted projects round-trip to the same project and Resource hashes, reject wrong passphrases and modified ciphertext, and leave no readable recovery/project payload behind.
- Cleanup PDF/JSON reports match applied/excluded/deferred findings, disclose target/rule versions and renderer, omit unrequested source content, and make no approval claim.
- Batch exports produce the expected separate deck/report files and never treat a subset as complete success.

### 18.6 MCP validation

- Tools unavailable when the app is closed.
- Non-loopback connection refused.
- Invalid or stale tokens refused.
- Unsafe connection-file permissions refused where supported.
- Read scopes enforced.
- Per-Resource and per-representation permission enforced and reset on restart and project close.
- Result sizes bounded and paginated.
- Unsupported multimedia representations fall back safely; raw originals, external paths, and package paths are never returned.
- Slide proposals may reference existing authorized Resource IDs but cannot mutate the Resource registry or bytes.
- Proposal writes are temporary until Apply.
- Apply/Reject behavior visible and undoable.
- Stale revision rejected.
- Design-thread anchors resolve to the exact stable object, text range, table cell/range, or normalized region; ambiguous remapping returns `needs-reanchor` rather than guessing.
- An AI reply alone cannot resolve a design thread; only an applied resolution or explicit human disposition can do so.
- Human direct edits and applied AI proposals enter the same validated transaction and undo/redo history.
- Operation-scope tool filtering enforced even against cached direct calls; audit has no mutation tools and cleanup cannot call reflow/generative tools.
- Cleanup rule proposals affect only explicit deck/object IDs and one allowed property under the pinned rule version.
- Design-thread and table proposals target exact stable IDs and current revisions, preserve protected cell/content hashes, and expose the rerendered result before Apply.
- No delete, export, publish, or storage mutation tool exposed in V1.

### 18.7 Desktop validation

- Build and typecheck.
- Unit and integration tests.
- Electron startup/smoke on macOS and Windows.
- Renderer sandbox/context-isolation assertions.
- Local network-policy assertions.
- Autosave and interrupted-write recovery.
- Standard-package readable-data warning/acknowledgement and encrypted-to-standard Save As boundary.
- Large-deck and large-Resource streaming, package-checkpoint, open, and responsiveness benchmarks.
- Resource intake cancellation, deduplication, dependency-aware removal, missing-external-original, and safe-package adversarial tests.
- Keyboard and focus checks.
- Representative desktop visual inspection.

### 18.8 Source-install validation

- Fresh anonymous clone on macOS.
- Fresh anonymous clone on Windows.
- Setup with MCP installation selected.
- Setup with MCP skipped.
- Rerun as update/repair.
- Dirty checkout refusal.
- Offline later launch after initial dependency/runtime setup.
- Exact revision and checksum confirmation.
- Installation PDF command/link verification.

---

## 19. Planned repository structure

```text
Presentation Studio/
├── assets/
│   └── templates/
│       └── 13_ORNL_Presentation_16x9_Template.potx
├── electron/
│   ├── main.cjs
│   ├── preload.cjs
│   └── network-policy.cjs
├── mcp/
│   ├── server.mjs
│   └── presentation-app-client.mjs
├── src/
│   ├── components/
│   ├── data/
│   ├── lib/
│   │   ├── project/
│   │   ├── package/
│   │   ├── resources/
│   │   ├── crypto/
│   │   ├── templates/
│   │   ├── import/
│   │   ├── batch/
│   │   ├── cleanup/
│   │   ├── rules/
│   │   ├── proposals/
│   │   ├── render/
│   │   └── export/
│   ├── workers/
│   ├── App.tsx
│   ├── main.tsx
│   ├── styles.css
│   └── types.ts
├── scripts/
│   ├── setup-macos.zsh
│   ├── setup-windows.ps1
│   ├── start-macos.zsh
│   ├── start-windows.ps1
│   ├── start-electron.mjs
│   ├── configure-presentation-mcp.mjs
│   ├── compile-template-pack.mjs
│   └── build-installation-guide.*
├── docs/
│   ├── Presentation-Studio-Installation-Guide.pdf
│   ├── PROJECT-FORMAT.md
│   ├── RESOURCES.md
│   ├── MCP.md
│   ├── TEMPLATE-PACKS.md
│   ├── CLEANUP-RULES.md
│   ├── BATCH-REVIEW.md
│   └── CONTENT-POLICIES.md
├── examples/
│   └── synthetic-public-fixtures-only/
├── tests/
├── Start-Presentation-Studio.command
├── Start-Presentation-Studio.cmd
├── PRESENTATION-STUDIO-SPEC.md
├── README.md
├── SECURITY.md
└── package.json
```

The production ORNL template is an authorized, versioned local application asset and is not committed to the public repository. Imported user content, extracted production Template Pack artwork, and standalone protected brand artwork remain outside the repository unless separately cleared.

---

## 20. Implementation phases and approval gates

### Phase 0 — Specification and repository safeguards

Deliverables:

- approved product/technical specification;
- repository data boundary for imported and generated content;
- ignore rules and tracked/history safety scan design;
- selected synthetic/public test content;
- local authorized-template installation design with version, provenance, and SHA-256 verification.

Gate: approve this specification, verify an authorized locally installed template against its local manifest, and confirm that repository safeguards exclude official template bytes plus imported and generated working content.

### Phase 1 — Project, Template Pack, and cleanup-rule foundation

Deliverables:

- single-deck/review-batch project schema and migrations;
- `.pstudio` package format, manifest, restricted path rules, and bounded streaming I/O;
- unified Resource registry, content-addressed blob storage, derivative lineage, and portability validator;
- encrypted-envelope schema, versioning, and cryptographic test vectors;
- Template Pack, Cleanup Rule Profile, finding, protection/exclusion, and style-exemplar schemas;
- local ORNL and project-local sponsor/custom template inspection/compiler;
- PowerPoint-native render bridge, renderer provenance, bounded render cache, and explicit qualified-fallback behavior;
- structural template classification with confidence/evidence;
- complete inventory of every layout in the active authorized ORNL revision (30 in the initial inspected reference);
- local versioned ORNL Template Pack installation, side-by-side update, active-version display, and rollback behavior;
- semantic placeholder mapping and compatibility rules for recreating content in any selected approved ORNL layout;
- synthetic committed Template Packs/rules for tests;
- deterministic validation.

Gate: compile and identify the locally installed authorized ORNL template plus a synthetic sponsor template without modifying either source; exercise every discovered ORNL layout with representative synthetic content and PowerPoint-native layout renders; create single-deck and review-batch self-contained packages, move/delete their external originals, and reopen them with the same deck, Resource, rule, exemplar, derivative, and Template Pack hashes. No official ORNL template or extracted asset may appear in Git history.

### Phase 2 — Cleanup-first PowerPoint import and audit vertical slice

Deliverables:

- branded Electron shell with **Clean existing PowerPoint** as the primary action;
- multi-file/folder Batch intake and immutable source records;
- layered OOXML preflight, native interpretation, preservation envelope, and visual baseline;
- explicit `editable-native`, `preserved-native`, `conversion-required`, and `unsupported-blocking` object states;
- modern-comment failure isolation and preservation;
- template classification/confirmation UI;
- effective-font, master/layout, local-override, table, figure, chart, notes, and unsupported-object inventories;
- read-only deck/batch audit with production/editorial/domain categories;
- hybrid slide scene/preservation model plus revision-bound native Source render and deterministic geometry observation;
- package Save/Open/Save As and desktop smoke tests.

Gate: audit the synthetic mixed-font ORNL, sponsor, table, and modern-comment fixtures without changing their source files. The comment-triggered high-level parser failure must be isolated while package preflight and preservation succeed; the sponsor deck must not receive ORNL findings. Every slide shows its native baseline when PowerPoint is available, and every object exposes an honest fidelity state.

### Phase 3 — Conservative cleanup, review, and editable PowerPoint export

Deliverables:

- versioned Cleanup Rule Profiles and confidence classification;
- Aptos/legacy-font mapping with inheritance, exemption, and fit-risk handling;
- Approved Style Exemplars for compatible tables and figures;
- approved-as-is/protected/excluded scopes and minimal-change validation;
- exact content/data/media/notes/comments hashing and transformation ledger;
- By Rule and By Slide Before/After review with automatic exception separation;
- automated Source → Current → Proposal native-render loop with deterministic and visual comparison evidence available to the AI;
- template-preserving editable PPTX export and source-versus-output audit;
- native PowerPoint versus alternate-renderer comparison;
- human-readable PDF and machine-readable JSON cleanup reports.

Gate: clean a synthetic multi-deck batch while preserving all protected content and prior-approved scopes, applying no ORNL rules to the sponsor deck, retaining native editable tables/figures, creating no fit regression, and producing separate PowerPoint copies/reports. Every proposal and output opens and renders in PowerPoint where available, and the AI receives the native comparison before the proposal can be marked ready for review.

### Phase 4 — Resumable 200-slide batch and MCP collaboration

Deliverables:

- responsive batch queue, per-deck checkpoints, pause/cancel/resume/retry, autosave, recovery, undo/redo, and failure isolation;
- standard and encrypted Save/Open/Save As with encrypted recovery;
- Canvas Observer with full-slide, crop, geometry, contact-sheet, and render-comparison views;
- AI slide work orders that combine native pixels, structured scene data, Template Pack candidates, protected-content hashes, deterministic findings, and allowed operations;
- local STDIO MCP server, loopback connection/token runtime, and per-Resource/per-representation permissions;
- cleanup audit/read tools and operation-scoped cleanup proposal tools;
- Before/After review, Apply/Reject/Exclude/Defer, stale-write behavior, and MCP install/remove configuration;
- measured 200-slide and 250-slide synthetic qualification runs on macOS and Windows.

Gate: at least two MCP-capable desktop clients inspect the same authorized batch/deck findings and slide revision, then stage the same bounded cleanup-rule proposal without receiving unapproved Resources or accessing reflow/generative tools. A 200-plus-slide batch completes or resumes after interruption with accurate per-deck status and no unauthorized content/template changes.

### Phase 5 — Collaborative canvas, table editor, reflow, and remaining exports

Deliverables:

- explicit Reflow/Hybrid scope changes and layout matching;
- separate cleanup/round-trip and composition export strategies behind the shared scene and command model;
- unsupported-object review decisions and controlled native passthrough;
- Canva-style manual slide/slot editing independent of AI, backed by stable scene IDs and validated command transactions;
- shared human/AI undoable history, revision guards, active-human soft locks, and AI affected-object activity indicators;
- anchored object/text/table-cell/region design threads with reference crops, reply/resolution lineage, reanchor handling, and scoped MCP tools;
- first-class native table editing, semantic roles/styles, cell-level content/structure hashes, constraint-based sizing, continuation-slide proposals, and independent export-render QA;
- PDF, SVG, and PNG export from the shared scene;
- full Resources workspace with preview, insert, where-used, package-size, and dependency-aware removal;
- actual-artifact overflow, placeholder, master/layout, editability, and cross-render validation.

Gate: a human can directly edit a supported slide and table without AI; an AI can inspect an exact location-anchored thread, stage a bounded visible repair, and retain the anchor after reflow or return `needs-reanchor` without guessing. Human edits and applied AI fixes share one reversible history. An explicitly authorized reflow changes layout without changing locked content, while a Cleanup job cannot access those operations. A merged-cell table and a dense continuation table preserve exact cell/structure hashes and remain native/editable after PPTX export. PPTX, PDF, SVG, and PNG each pass their own inspected-artifact checks.

### Phase 6 — New presentations from source materials

Deliverables:

- DOCX, PDF, Markdown, text, CSV, and Excel intake plus the tested multimedia support matrix;
- bounded local extraction, Resource search, derivatives, locator registry, and citations;
- outline review and assertion-evidence slide composition;
- source-grounding validation and `[Sources]` notes;
- vetted local icon catalog, safe SVG/native-shape resolution, and Visual Need review;
- optional quarantined generated-resource candidate intake;
- generative MCP tools available only under Compose/Hybrid authorization.

Gate: a cleared source document produces a reviewable deck whose assertions/evidence trace to exact source locators, whose exported notes include `[Sources]` blocks, and whose generated visuals remain purpose-driven, authorized, and provenance-complete.

### Phase 7 — Source setup and installation PDF

Deliverables:

- macOS setup/start/update scripts;
- Windows setup/start/update scripts;
- optional MCP configuration;
- verified installation-guide PDF;
- fresh-clone validation evidence on both platforms.

Gate: both platform workflows pass from a fresh anonymous clone, the documented commands match the scripts, and no installer or security-control bypass is produced.

---

## 21. Release acceptance criteria

A Presentation Studio release candidate is not ready until all of the following are true:

1. The package format, project schema, Resource registry, Template Pack, and migrations pass deterministic tests.
2. The application starts through the source setup on supported macOS and Windows systems.
3. The Electron security and loopback-only MCP boundaries are verified.
4. Manual editing works without an AI client.
5. MCP changes remain review proposals until a person applies them.
6. All Resource-content access is off by default and resets off on restart and project close.
7. Compose mode preserves source traceability and does not invent unsupported claims.
8. Reflow mode makes no unauthorized locked-content changes.
9. Unsupported PowerPoint objects are preserved or explicitly resolved, never silently flattened.
10. PPTX, PDF, SVG, PNG, `.pstudio`, encrypted `.pstudio-secure`, and content-only JSON outputs each pass their own artifact checks.
11. Exported PPTX retains the intended master/layout structure, source notes, and supported native editability.
12. Every final slide render is inspected for overflow, overlap, crop, contrast, font substitution, and template drift.
13. The installation PDF matches independently tested setup commands.
14. No standalone protected logo, imported source, user project, export, credential, or runtime token is present in the public Git index or reachable history without explicit clearance; the authorized official template matches its expected hash.
15. Release notes call the product and ORNL outputs drafts unless formal approval has been documented.
16. Encrypted projects authenticate before opening, reject wrong passphrases or tampering, preserve project/Resource hashes, use encrypted recovery state, and never imply that external originals or derived exports are protected.
17. A standard or encrypted project moved to another qualified machine reopens and renders without access to any imported file's original path; every required Resource and Template Pack member resolves inside the package.
18. MCP lists and reads only explicitly authorized Resource representations, uses bounded multimedia fallbacks, never returns external/package paths or unapproved original bytes, and can stage slides that reference existing Resource IDs without mutating the Resource library.
19. Standard saves and encrypted-to-standard conversions show and record the required readable-data acknowledgement before packaging unencrypted Resource content.
20. PowerPoint package preflight and preservation succeed even when a valid but unfamiliar comment, metadata, or object part causes a high-level parser failure; the failure is isolated and reported.
21. Cleanup and strict reflow prove exact preservation of protected text, tables, chart data, notes, media bindings, comments, slide order, and hidden state at the granularity defined by policy.
22. An authorized AI client can inspect a revision-bound slide render plus structured geometry, stage a targeted change by stable object ID, and inspect the rerendered proposal without applying it.
23. New icons and generated visuals have a stated communication purpose, provenance, rights state, alt text, and human acceptance; CSS/web-font icons and unsafe SVG never reach export unresolved.
24. Final PPTX qualification compares a PowerPoint-native render with the approved project render where PowerPoint is available and reports renderer divergence rather than hiding it.
25. Imported decks default to `cleanup-only`; wording changes, new visuals, slide splitting, layout replacement, and cross-template conversion remain unavailable until a person changes scope.
26. Every deck has a confirmed template classification/target before cleanup, and sponsor/custom decks receive no unintended ORNL-specific rules or assets.
27. The font audit resolves effective fonts through the complete inheritance chain, protects declared equation/symbol/sponsor exemptions, and rerenders every proposed mapping for fit changes.
28. Approved table/figure exemplars apply only to compatible objects and preserve exact data, merged structure, media/chart identity, and semantic color roles.
29. Previously cleaned or `Approved as-is` slides/objects remain unchanged; exclusions and rejected findings remain durable within the project.
30. Production blockers, editorial-consistency questions, and domain technical-review questions remain separately labeled and never produce a false “technically correct” or approved state.
31. A qualified 200-plus-slide review batch supports pause/resume and per-deck failure isolation, exports separate non-overwriting PPTX copies/reports, and reports partial versus complete status accurately.
32. Cleanup reports match the actual applied/excluded/deferred transformations, rule/template versions, unresolved findings, source/output hashes, and renderer evidence without reproducing unrequested source content.
33. When a prior cleaned revision is supplied, three-way comparison can recover selected approved formatting without rolling current text, data, notes, media, ordering, or new objects back to the earlier deck.
34. A user can directly select and edit supported slide objects and native tables without MCP; each action is validated, undoable, and represented identically in the canonical scene and editable PowerPoint export.
35. A design thread can target an exact object, text range, table cell/range, or slide region; its pin follows an unambiguous reflow, retains original revision/crop evidence, and becomes `needs-reanchor` rather than attaching to the wrong object.
36. An authorized AI client can stage and reply with a bounded design-thread resolution against the current revision, but cannot mark the thread resolved, apply the proposal, or overwrite a concurrent human edit.
37. Table QA measures every cell and verifies text fit, type size, padding, alignment, merge topology, borders, contrast, semantic color, balanced geometry, and exact content/structure hashes in both project and exported PowerPoint renders.
38. Source, Current, Proposal, and Export renders expose renderer provenance and exact revision/package hashes; an approximate render cannot satisfy a native-review requirement.
39. Every imported object has an explicit fidelity state, and no unsupported object is silently flattened, omitted, duplicated, or presented as fully editable.
40. When PowerPoint is available, an AI design proposal is materialized and natively rendered before it can be marked ready for review; the exported artifact is rendered and compared again before completion.
41. A private, hash-pinned prior cleaned deck or approved equivalent can act as a local visual-regression fixture without adding its deck bytes, slide renders, or production Template Pack assets to Git.

---

## 22. Risks and decisions to resolve

### 22.1 Official template integrity and versioning

The official `.potx` is authorized for local Presentation Studio use but is not committed to the public repository. The locally installed source must remain byte-identical to the approved source for its recorded version. Template updates require a new expected hash, complete layout inventory, compatibility review, and an explicit project migration proposal; they must not silently change existing projects.

### 22.2 PowerPoint import/export library boundary

PptxGenJS supports native generation but is not by itself a complete imported-template round-trip solution. The implementation needs a tested template/OOXML preservation layer. This must be proven before claiming high-fidelity export.

### 22.3 Unsupported PowerPoint features

SmartArt, equations, animation, OLE, macros, 3D, embedded media, and specialized charts require explicit support/preservation decisions. The product must favor visible limitation reporting over lossy hidden conversion.

### 22.4 Exact text versus improved assertions

Strict reflow cannot create assertion-evidence headlines without changing content. The UX must keep assertion suggestions separate until the user changes the relevant policy.

### 22.5 Large-deck resource use

ZIP parsing, hashing, image extraction, preview rendering, large media copying, package checkpointing, encryption, and export can be memory- and I/O-intensive. Worker isolation, bounded streams, batching, cancellation, unchanged-member reuse, visible save states, and measured platform limits are required before publishing maximum-size claims. A self-contained package intentionally trades smaller files for portability; the Resources UI must show that cost before large imports.

### 22.6 Model variability

MCP compatibility does not make model outputs equivalent. Authoritative validation, exact schemas, bounded tools, source references, and human review must compensate for model-specific planning and writing differences.

### 22.7 Sensitive source material

Presentation Studio can enforce a local default and explicit read permissions, but it cannot authorize a remote model for sensitive ORNL content. Users remain responsible for selecting an approved environment. The application and guide must communicate this without suggesting that branding makes content safe to disclose.

### 22.8 Encrypted package performance and passphrase custody

Large presentation assets make whole-file, Base64, and all-in-memory encryption impractical. The encrypted container needs bounded or streaming authenticated processing, atomic failure behavior, and large-project benchmarks. The passphrase is not recoverable; users must keep it separately in an approved password manager. Encryption protects the saved package and encrypted recovery state, but it does not replace endpoint security, approved storage, access controls, retention policy, or verified backups.

### 22.9 Resource safety and active content

Documents, spreadsheets, presentations, SVGs, archives, audio, and video may contain malformed, oversized, deceptive, or active content. Intake must use signature validation, restricted parsers, decompression limits, no script/macro execution, safe previews, and explicit unsupported states. Preserving an original Resource does not authorize Presentation Studio to execute or fully interpret it.

### 22.10 Multimedia and model capability variance

Some MCP clients can consume text and images but not audio, video, large tables, or the standard MCP Resource capability. Presentation Studio must advertise capabilities, provide transcripts/poster frames/structured chunks when available, and show what was actually shared. “MCP-compatible” must not be represented as equivalent multimedia understanding across every model.

### 22.11 Resource and template rights

Self-containment does not create redistribution rights. The project records rights/authorization metadata and embeds fonts, template sources, stock media, or other licensed material only when allowed. The application must warn before packaging known restricted material but cannot determine legal clearance automatically.

### 22.12 Parser tolerance and round-trip preservation

Office files can contain valid parts that a selected library does not recognize, including newer comment schemas. Treating one parser as authoritative could reject a usable deck or silently discard content. The architecture therefore requires raw-package preflight, per-part failure isolation, preserved source parts, synthetic regression fixtures, and an exported-package audit. Round-trip fidelity remains a measured capability by object type, not a blanket promise.

### 22.13 Visual-model confidence and renderer divergence

A model may miss subtle alignment, crop, legibility, or consistency problems even when it receives a slide image, and two presentation renderers may disagree. Canvas observation must combine pixels with exact geometry and deterministic checks. Final quality claims are based on the inspected exported artifact, with PowerPoint-native rendering used where available and every known divergence disclosed.

The native render bridge is an optional local integration rather than a reason to weaken offline behavior. PowerPoint version differences, platform automation restrictions, foreground-window requirements, missing fonts, protected-view prompts, and add-in/API limitations must be detected and reported. Presentation Studio retains a qualified fallback path, but a fallback does not inherit native authority merely because PowerPoint is unavailable.

### 22.14 Generated-visual safety and factual implication

Generated imagery can disclose sensitive source material, introduce unsupported facts, imitate protected content, or create attractive but irrelevant decoration. Visual Needs require a communication purpose, minimal authorized context, provenance, rights acknowledgement, and a human gate. The application does not claim that automated checking proves factual, legal, or brand approval.

### 22.15 Intentional edits versus inconsistency

A previous editor or customer may have made a deliberate local change that resembles accidental legacy formatting. Frequency and deviation from a template are evidence, not proof. The product mitigates this with inheritance/source inspection, style clustering, approved-as-is protection, confidence classes, minimum-property proposals, durable exclusions, and human review. It cannot infer creative or technical intent perfectly.

### 22.16 Font-metric and platform drift

Century Gothic, Arial, Aptos, sponsor fonts, substitutions, and different Office/platform renderers have different metrics. A semantically safe font-family change can alter line breaks or overflow. Every mapping therefore requires effective-font resolution, exemption handling, rerendering, and final PowerPoint inspection. A batch font replacement is not safe merely because text hashes match.

### 22.17 Exemplar authority and semantic styling

A sample table or figure may encode one editor's preferred treatment, a customer-specific exception, or semantic colors that should not transfer. Exemplar approval must state its owner, scope, permitted properties, and semantic exceptions. Structural incompatibility or unclear authority produces a review finding rather than forced normalization.

### 22.18 Sponsor-template scope and rights

Sponsor templates may contain protected artwork, licensed fonts, or rules that differ from ORNL guidance. Project-local packs must preserve their source/rights metadata, avoid repository/global installation by default, and prevent cross-template contamination. The user remains responsible for authorization to use and package sponsor assets.

### 22.19 Ambiguous technical-review expectations

“Technical checks” can mean mechanical production QA, editorial consistency, or domain correctness. Conflating them creates false confidence. Each batch needs a visible checklist and responsible reviewer; domain correctness remains unresolved unless supported by authorized sources and a qualified human decision.

---

## 23. Initial product definition

Presentation Studio is defined as:

> A local-first, template-aware Electron presentation production system whose first job is to audit, visually improve, and conservatively clean batches of existing ORNL or sponsor PowerPoint decks without rewriting content, erasing intentional prior work, or applying the wrong template. It stores single decks or review batches as self-contained packages with a canonical editable scene, a native OOXML preservation envelope, pinned Template Packs, and embedded Resources; gives compatible MCP clients native-rendered pixels plus exact structured design context; iterates through bounded human-reviewable proposals; exports separate editable PowerPoint copies plus reports/PDF/SVG/PNG; and later uses the same platform to compose source-grounded assertion-evidence presentations.

The first visual expression is balanced and ORNL-aligned, while sponsor decks retain their own template expression. The first distribution is source-only. The first AI boundary is local MCP with explicit per-Resource permissions, operation-scoped tools, and review-first writes. The first proof of quality is a preserved source, an explainable audit, a native-rendered design iteration, and an inspected editable PowerPoint copy—not merely a clean app canvas.
