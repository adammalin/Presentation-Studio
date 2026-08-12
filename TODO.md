# Presentation Studio TODO

- **Development-plan version:** 0.9
- **Architecture focus:** PowerPoint-native visual evidence, lossless imported-content preservation, and an iterative human/AI ORNL design loop.

## Authorized ORNL Template Pack

- [ ] **Install, version, and qualify the latest authorized ORNL PowerPoint template locally.**
  - **Implemented foundation (2026-08-11):** local POTX/PPTX selection, independent SHA-256 verification, protected app-data storage, restart persistence, full 30-layout inventory for the current ORNL reference, real template-media reuse, searchable/category-filtered gallery, selectable detail view, and app-only placeholder guides.
  - **Implemented semantic/native slice (2026-08-12):** duplicate inherited placeholders are normalized into semantic title, label, text, image, table, chart, media, and flexible-content slots; every layout now exposes intent, capacity, readable-type constraints, and deterministic compatibility scoring. The app materializes one temporary slide per real custom layout, renders all 30 through Microsoft PowerPoint, overlays Studio-only semantic guides, caches the native results for the session, and exposes both the full catalog and per-imported-slide recommendations through model-independent MCP tools. The authorized local source and extracted artwork remain outside Git.
  - **Still required:** authoritative revision/effective-date/owner metadata, side-by-side pack selection and rollback, project-pinned pack snapshots, staged semantic reflow, representative synthetic content filling/overflow qualification for every layout, disk-cache policy, and Windows native qualification.
  - Import the approved `.potx`/`.pptx` from a local file or managed local folder; do not bundle it in or commit it to the public repository.
  - Record the template name, revision/version, effective date when known, source/owner note, import date, SHA-256, and compatibility status in a local manifest.
  - Keep Template Pack versions side by side, show the active version in the UI, preserve project-pinned versions, and require a reviewable migration rather than silently replacing a template.
  - Inventory every real master and layout in the installed revision and generate a layout gallery from the actual template.
  - Map semantic content roles to native placeholders so a person or MCP client can stage recreation/reflow into any selected approved ORNL design.
  - Preserve exact text, data, notes, media identity, and editable PowerPoint objects in locked/reflow-only work; report incompatible layouts, unmapped content, and overflow instead of forcing a fit.
  - Keep sponsor/custom Template Packs isolated and require an explicit target-template decision before cross-template conversion.

### Acceptance criteria

- The active authorized template version, provenance, and SHA-256 are visible and auditable.
- Every master/layout discovered in the installed ORNL template appears in the gallery and can be selected for a staged proposal.
- Representative synthetic content is compiled and rendered through PowerPoint for every discovered layout, with placeholder, theme, master, overflow, and editable-PowerPoint checks.
- Reopening a project uses its pinned Template Pack snapshot even after another authorized revision is installed.
- Repository and history scans find no official ORNL template, extracted template artwork, or production Template Pack bytes.

## Product roadmap

### Round goal and implementation order

The current implementation provides a useful OOXML inspection and bounded geometry foundation. Continue in this order so the AI designs against trustworthy evidence before the editor gains broader authoring features:

1. build the local PowerPoint-native render bridge and make renderer provenance visible;
2. complete the ORNL Template Pack compiler with native thumbnails, semantic slots, constraints, and compatibility scoring;
3. formalize the hybrid scene graph plus OOXML preservation envelope and object fidelity states;
4. close the AI design loop across Source → Current → Proposal → Export renders;
5. implement semantic layout selection, the cleanup/round-trip engine, and the composition engine;
6. build the first-class native table solver and editor;
7. expand the collaborative canvas, anchored comments, multi-select, guides, undo/redo, cropping, and property editing;
8. add source-grounded assertion-evidence composition, multimedia workflows, remaining exports, and large-batch qualification.

### Current qualification evidence

- **Passed live bounded-design test (2026-08-12):** a 26-slide local deck was inspected through the live STDIO MCP, one exact-content cover alignment was staged as an atomic slide transaction, reviewed in Current/Proposal, accepted, and exported as a new editable PPTX. Microsoft PowerPoint independently rendered all 26 source and output slides at 1600 × 900. Only the intended first-slide JPEG changed; the other 25 were byte-identical. Slide count, slide/text-box text hashes, table-cell hashes, and merged topology all matched.
- This proves the bounded object-layout path and stale-write protection. It does not complete deck-wide autonomous reflow, semantic layout selection, icon/image composition, native table solving, or an in-app PowerPoint render/compare gate.

- [ ] **Build the local PowerPoint-native render bridge and provenance-aware render cache.**
  - **Implemented foundation (2026-08-12):** the macOS Electron bridge now copies an exact temporary PPTX revision into PowerPoint's local sandbox, exports it through Microsoft PowerPoint as PDF, rasterizes bounded 1400-pixel JPEGs locally, verifies slide-image dimensions/count, deletes the cleartext render job, serializes concurrent jobs, and exposes authoritative renderer provenance to Slides, Current/Proposal Review, Designs, and MCP `source|current|proposal|export` renders. The OOXML renderer is now visibly labeled as an approximate fallback. The installed ORNL reference successfully materialized and rendered all 30 custom layouts at 1400 × 788 with no warnings. PowerPoint version and raster hashes are captured in memory; persistent disk-cache policy, selected-object crops/contact sheets, Windows automation, exported-file rerender comparison, and permission-recovery UX remain.
  - Render exact Source, Current, Proposal, and Export revisions through Microsoft PowerPoint when it is installed without modifying the imported original.
  - Materialize Current and Proposal as temporary candidate PPTX files from exact package/scene revisions, render them locally, and discard temporary cleartext artifacts according to the project's encryption and recovery policy.
  - Record renderer family/version, authority level, package hash, scene revision, dimensions, slide index, font substitutions when available, raster hash, and failure diagnostics for every render.
  - Support `powerpoint-native`, `qualified-alternate`, and `studio-approximate` authority labels. Never show an approximate preview without its label or allow it to satisfy a native-review gate.
  - Feed the same native slide image, selected-object crop, deck contact sheet, and render metadata to the Slides UI, Review UI, deterministic checks, and authorized MCP clients.
  - Keep the current OOXML renderer as a diagnostic/editor fallback; do not treat it as the final visual source of truth.

#### Acceptance criteria

- The Slides tab can display PowerPoint-native thumbnails and a selectable full-slide render for every slide in a supported local deck.
- A candidate proposal is automatically materialized and rerendered without overwriting the source deck.
- Source, Current, Proposal, and Export views identify their exact revision and renderer; stale or approximate renders cannot masquerade as current native evidence.
- A PowerPoint-unavailable machine can continue with an explicit fallback state, but completion remains `Needs native review`.

- [ ] **Formalize the hybrid scene graph and PowerPoint preservation envelope.**
  - **Implemented foundation (2026-08-12):** every newly audited deck now compiles into a versioned `presentation-studio/scene` stored in the project. All directly inventoried slide objects retain stable import IDs, native geometry, z-order, semantic role, source slide/shape locators, content hashes where available, one visible fidelity state, represented-property limits, and an explicit allowed-operation matrix. Every source slide XML part and relationship part is SHA-256 bound into a `presentation-studio/preservation-envelope`; original embedded PPTX bytes and PowerPoint-native renders remain authoritative. The Slides object editor shows fidelity state/reason and locks unsupported operations; Deck Audit summarizes scene coverage/blockers; MCP exposes content-minimized scene summary and per-slide object tools. Legacy projects reopen without inventing a scene and rebuild from their embedded source. The real 26-slide EMT source qualified at 120 objects, 26/26 hashed slide parts, 99 editable-native objects, 21 preserved-native objects, and no package blockers. Scene-driven proposal materialization, part-level byte-identity regression fixtures, precise SmartArt/OLE adapters, transaction history, and composition-scene export remain.
  - Make the canonical scene authoritative only for supported editable properties, semantic roles, template bindings, comments, and transactions.
  - Retain source OOXML parts and relationships for objects that are not fully represented, keyed to stable import IDs and protected by package/content hashes.
  - Assign every imported object one visible fidelity state: `editable-native`, `preserved-native`, `conversion-required`, or `unsupported-blocking`.
  - Permit safe whole-object positioning for qualified preserved objects while keeping unsupported internal properties locked.
  - Merge scene edits with preserved package content during cleanup export; never silently flatten, approximate, omit, or duplicate preserved objects.
  - Keep cleanup/round-trip export separate from new-slide composition export even though both use the shared scene and command model.

#### Acceptance criteria

- Import, save/reopen, proposal materialization, and export retain every supported or preserved object with the same stable identity or a documented mapping.
- Unsupported content remains visible through its native baseline and is not replaced by a placeholder in review or export.
- A lossy conversion cannot occur without an explicit, reviewable `conversion-required` decision.
- Round-trip regression fixtures prove that untouched slide/package parts remain byte-identical where the package contract permits it.

- [ ] **Close the AI visual-design loop and qualify it against a private golden deck.**
  - **Implemented first iterative slice (2026-08-12):** MCP can now build one revision-bound slide work order—or a representative five-slide deck set—combining authoritative Current PowerPoint raster identity/provenance, exact locked text, the hybrid scene, allowed object operations, findings, submitted design comments, ORNL rules, and ranked semantic Template Pack candidates. An AI can bind source objects to approved semantic zones, materialize the exact-content geometry proposal, request a side-by-side Current/Proposal Microsoft PowerPoint comparison with raster hashes and changed-region metrics, and reject its own failed pending draft with the native evidence and concrete rationale attached. Existing text frames have a safer horizontal-zone mode, while substantial unmeasured vertical/shrink replacement is rejected. A guarded native-layout operation now reuses a byte-exact approved layout already carried by the deck or clones the allowlisted native dependency graph, maps compatible placeholder identities, and repoints only one slide. Work orders expose text-frame insets, direct paragraph margins/indents, bullets, confidence, and estimated optical text starts so alignment decisions use visible text rather than box edges. The real EMT slide 20 retained all visible-text hashes and rendered successfully through PowerPoint after exact layout reuse; an earlier cloned-master attempt lost visible hierarchy/logo relationships and was correctly treated as failed qualification evidence rather than accepted.
  - **Implemented atomic native-layout slice (2026-08-12):** `stage_slide_recomposition` now performs one reversible transaction containing the approved native layout remap plus all validated semantic object bindings. On the real EMT slide 20, the one-column ORNL layout, title/body placeholder mapping, and two source-bound horizontal placements materialized together; every slide/text hash stayed exact and PowerPoint retained the ORNL logo, footer, bullets, hierarchy, and editable text while producing a cleaner common grid.
  - **Still required:** qualify/fix cloned-master artwork across representative layouts, high-level distribution and measured text-fit solvers, table/figure/icon operations, bounded automatic revision orchestration, approved-as-is evidence, deck contact sheets, exported-file comparison, and the private hash-pinned golden-deck gate.
  - Give the model one versioned work order containing native pixels, structured objects, exact content protections, ORNL rules, Template Pack layout candidates, authorized Resources, deterministic findings, and allowed operations.
  - Implement the autonomous loop: Inspect → Diagnose → Propose → Materialize candidate PPTX → Native render → Compare → Revise → Review.
  - Add high-level MCP operations for layout recommendation/application, slide recomposition, alignment/distribution, text fitting, table design, figure/caption treatment, semantic icon insertion, visual-need creation, and visual preflight.
  - Require an explicit per-slide disposition and a deck contact-sheet pass; do not stop after font normalization or the easiest geometry changes.
  - Register an approved prior cleaned deck as a private, hash-pinned local visual-regression fixture. Store no client deck, extracted slide image, or production template byte in Git.
  - Automatically iterate routine reversible proposals up to a bounded cap. Pause only for template ambiguity, protected-content conflict, technical meaning, unsafe disclosure, unsupported/lossy conversion, or final external export.

#### Acceptance criteria

- An authorized AI can inspect, redesign, rerender, and reassess a slide without guessing from approximate geometry or asking the user about routine design choices.
- Every in-scope slide is improved or recorded as `approved-as-is` with native-render evidence and deterministic findings.
- The private golden-deck comparison detects regressions in hierarchy, layout balance, table quality, template fidelity, clipping, and editability before a change is accepted.
- No AI-ready claim is produced until required native-render and exported-artifact checks pass.

- [ ] **Ship one versioned Presentation Design Standard as the app, preflight, and MCP source of truth.**
  - **Implemented foundation (2026-08-12):** one shared versioned JSON standard now drives new-project defaults, resolved ORNL deck profiles, the Rules view, table fallback tokens, geometry/fit policy, drift tests, and MCP cleanup/design-contract responses. Ordinary current/older ORNL imports adopt 16:9 + Aptos + exact-content cleanup without a target-template questionnaire. Schema negotiation, generated preflight views, and full cleanup-engine enforcement remain.
  - Store the standard as a machine-readable, versioned local ruleset rather than relying on prompt prose. Generate the human preflight checklist and the MCP `get_design_contract` response from that same ruleset so the UI, cleanup engine, tests, reports, and every compatible AI model receive identical instructions.
  - Provide an **ORNL presentation** quick-start that is selected automatically for a new project unless the imported deck is confidently classified as sponsor/custom or the user explicitly chooses another profile.
  - Default to a 16:9 canvas, the active authorized ORNL Template Pack, Aptos presentation typography, square brand-created containers, exact-content preservation, editable native PowerPoint output, and restrained ORNL Green/Hale Navy/neutral color roles unless a supplied template or explicit user instruction overrides them.
  - Treat the active authorized Template Pack and an approved project exemplar as higher authority than generic fallback rules. Record every override in the project profile so the same deck renders consistently after reopen and on another supported machine.
  - Include the ORNL presentation baselines in the contract: assertion-led hierarchy when content editing is allowed; body text at least 16 pt; captions/labels at least 14 pt; approved palette roles; consistent icon family/stroke; relevant, cleared imagery; preserved master/layout/theme; and exported-artifact visual inspection. Do not represent the product defaults as formal ORNL approval.
  - Make **Clean up with ORNL defaults** the primary action. It should start the audit/design pass without a setup questionnaire and should preserve the source as read-only while writing a new draft/output.
  - Keep routine choices autonomous: layout selection, alignment, spacing, safe margins, font normalization, nonsemantic color cleanup, table/figure treatment, and fit-safe sizing should not each create an approval. Pause only for a genuinely ambiguous target template, protected-content conflict, unsupported object, technical/semantic ambiguity, unsafe Resource disclosure, or external save/publish action.
  - Show a compact summary of the adopted defaults with one **Customize** path; do not require the user to confirm defaults they did not change.
  - Add schema/version negotiation to MCP. A client with an unknown contract version receives a readable fallback plus an explicit compatibility warning, never a silent partial ruleset.

#### Acceptance criteria

- A first-time user can import an ordinary ORNL deck and start a complete exact-content cleanup with one primary action and no preliminary approval sequence.
- New projects are 16:9 and Aptos by default; imported sponsor/custom decks are not silently ORNL-restyled.
- The UI profile summary, generated preflight document, cleanup engine, MCP design contract, cleanup report, and tests all expose the same standard version and resolved values.
- Automated drift tests fail when a table, typography, color, geometry, or autonomy rule differs between the machine-readable standard and an MCP/UI representation.
- Only exception decisions interrupt a run; ordinary deterministic design work proceeds and appears in the proposal/review state.

- [ ] **Define and enforce a deterministic ORNL native-table style profile.**
  - **Implemented foundation (2026-08-12):** the versioned standard now contains concrete native/editable Aptos type, header/body color, banding, stroke, padding, minimum-size, preservation, and overflow tokens; the Rules UI and MCP expose the same values. Designer Cleanup now applies those tokens to compatible native tables, shows Current/Proposal comparisons, records dense/complex/semantic-color exceptions, and rejects any export that changes exact cell content or merged topology. Semantic-role editing, cell-level findings, measured native-PowerPoint overflow, and independent export QA remain.
  - Add role-based tokens for table title/caption, header, body, first/last column, subtotal/total, footnote/source, semantic highlight, continuation header, fills, strokes, banding, type, alignment, cell margins/padding, and minimum row height. Apply the same resolved tokens to every compatible table in the deck.
  - Use this product fallback when the active authorized Template Pack or approved exemplar does not provide a more specific table treatment:
    - square native PowerPoint table; Aptos throughout;
    - Hale Navy header fill with Polar white, bold header text;
    - Dark Matter body text on Polar white, with restrained Graphite row banding when it improves scanning;
    - minimal rules rather than a boxed spreadsheet grid: consistent Graphite horizontal separators, no decorative colored outer outline, and vertical rules only when required to preserve column comprehension;
    - 16 pt header and body text by default, with captions/labels no smaller than 14 pt; dense tables must reflow, widen, or continue rather than silently shrinking below the resolved minimum;
    - consistent 6 pt left/right and 4 pt top/bottom cell padding, then measured fit using the final exported font metrics;
    - left-align prose, align numbers consistently by meaning, center only short categorical headings, and vertically center headers while keeping multiline body content readable;
    - preserve all semantic colors, exact cell text, row/column order, merged topology, units, emphasis carrying meaning, notes, and source/caption relationships.
  - Treat the values above as versioned Presentation Studio defaults, not as a claim that every listed value is an official ORNL mandate. A qualified Template Pack rule or approved exemplar may replace them visibly and project-wide.
  - Compare tables by semantic role, not only by source formatting. Report intentional exceptions and structurally incompatible tables instead of forcing a misleading match.
  - Add a deck-wide **Normalize compatible tables** proposal with a live Before/After table crop and an exceptions list. One proposal may cover deterministic-compatible tables; it must not require a separate approval for every table.
  - Preflight native tables for inconsistent font family/size/weight, header/body role drift, fills, borders, line weight, padding, alignment, banding phase, row height, column balance, wrapping, clipping, overflow, contrast, caption/source separation, and alignment with neighboring objects.

#### Acceptance criteria

- Two compatible tables with the same semantic roles resolve to identical table tokens in project state and editable PPTX export.
- Mixed source outlines, padding, fonts, heading weights, fills, and banding are normalized or identified as explicit semantic exceptions.
- The exported deck contains no clipped/missing table text, accidental border variation, font below the resolved minimum, or unsupported table rasterization.
- PowerPoint-native and independent renders are compared after export; a material difference blocks completion and links to the exact table/cell/edge.
- The MCP contract exposes resolved table tokens, exceptions, stable cell IDs, and the required preservation/QA rules to any compatible AI model.

- [ ] **Implement the full Designer Cleanup engine and visual review canvas.**
  - **Implemented foundation (2026-08-12):** the Slides tab renders every current embedded PPTX slide locally from OOXML, inherited layout geometry, editable text/shapes, images, connectors, and native table cells; thumbnails open a large revision-labeled review surface. Designer Cleanup now inventories every directly editable text box, estimates fit from final box geometry and direct run size, records overflow/off-slide/safe-margin/ambiguous-alignment exceptions, repairs collision-checked high-confidence cover and repeated peer-edge drift, assigns change/approved-as-is/needs-review dispositions, materializes selected changes in memory, and shows side-by-side Current/Proposal designs before acceptance. Review now provides persistent per-slide Approved/Changes requested decisions, one-click Approve & next, Approve all & continue, a progress rail, a focused revision-request composer, and direct Edit/Point comment handoffs that reopen the exact proposed PowerPoint-native slide; advanced command-group selection is intentionally collapsed. An AI revision reopens only its affected slide while preserving other approvals, and accepted designs remain the editable proposal baseline. The selected slide's geometry checks are visible in Review, audit reports omit text while retaining geometry findings, and MCP receives structured text-box geometry plus both bounded renders. Optical alignment, overlap checks across all object types, authoritative font measurement, Source/Current semantic distinctions, Export comparison, zoom/filter context, layout remapping, full geometry reflow, and independent PowerPoint export rendering remain.
  - Treat an explicit request to improve every slide as `reflow` plus exact-content `reflow-only`, after one target-template decision.
  - Evaluate every slide and every text box for alignment, safe-area placement, overflow, wrapping, insets, typeface, size, weight, color, contrast, line spacing, and reading order.
  - Choose compatible approved layouts, normalize tables/figures/captions, and use semantic icons or native diagrams only when they clarify existing content.
  - Render source/current/proposal/export representations, show the actual current design of every slide in the Slides view, and require independent exported-artifact contact-sheet inspection before completion.
  - Keep routine design decisions autonomous; request one focused approval only for content fit, technical ambiguity, target-template ambiguity, unsupported objects, or external save/publish boundaries.
  - Make each slide thumbnail selectable. A click opens a large, zoomable, revision-labeled slide viewer with Source/Current/Proposal/Export comparison, findings, comments, and a direct path into Edit or Comment mode.
  - Keep thumbnail and close-up renders revision-bound, invalidate them after a change, and never show a structural placeholder or stale source render as the current design.

### Acceptance criteria

- Every in-scope slide has a visual-review disposition and no slide is silently skipped after font cleanup.
- Every exported slide passes clipping, overflow, off-slide, accidental-overlap, readability, and content-preservation checks in an independent renderer.
- The app shows actual source and proposed slide designs before export rather than structural placeholder cards.
- Selecting a Slides-tab thumbnail opens the exact current revision at a useful review size, and returning to the gallery preserves the user's slide, zoom, and filter context.
- MCP clients receive the versioned Designer Contract and can request exact bounded slide design context without receiving unapproved Resource bytes.

- [ ] **Build the shared human/AI slide editor and anchored design-review workflow.**
  - **Implemented foundation (2026-08-12):** Comment mode supports click/drag normalized-region anchors on the current or proposed slide, private saved notes, submitted AI threads, revision binding, project-package persistence, and bounded MCP list/read tools that exclude private notes. Submitted comments from Review mark only their slide for revision without making the proposal stale; private notes preserve the current decision. The audit now assigns stable slide-local object IDs to supported text, shape, picture, table, chart, connector, group, and graphic-frame objects. Edit mode overlays those objects on the real slide preview and supports selection, snapped drag, keyboard-like nudges, measured position/size fields, and safe-area alignment. Human and MCP geometry edits share one validated proposal command, reject stale/unknown/protected/off-slide targets, preserve source bytes and exact text/table hashes, and appear in Current/Proposal before Apply. MCP clients can now read exact object-mapped text and inch geometry, then stage 1–20 related moves/resizes as one atomic slide transaction; the validator rejects new overlap, worsened fit, unsafe-edge regression, and accidental picture distortion unless a deliberate exception is recorded. Multi-select, rotation/cropping, smart guides, object-bound comment anchors, transaction history/undo, reference crops, replies, resolution states, and soft locks remain.
  - Replace the structural Slides placeholder with a revision-bound direct-manipulation canvas: select, multi-select, edit text in place, move, resize, supported rotate, crop, snap, smart guides, align, distribute, group, lock, layers, keyboard nudge, and accessible undo/redo.
  - Route human edits and applied AI proposals through one stable-ID command/transaction model with validation, revision guards, coalesced pointer history, and short-lived human-edit soft locks.
  - Add Edit and Comment modes. Comments must anchor to a stable object, text range, table cell/row/column/range, or normalized region and retain original revision/render/crop fallback evidence.
  - In Comment mode, let the user click an element or drag a precise region, see the anchor highlight/crop, write the requested adjustment, and choose **Submit to AI** or **Save note** without leaving the slide viewer.
  - Add threaded replies, open/resolved/reopened/needs-reanchor states, assignee/filtering, affected-object highlighting, and Apply/Reject/Revise/Resolve/Reopen controls.
  - Add MCP read/proposal tools for bounded design threads. An AI reply cannot resolve or apply a thread, and ambiguous remapping must return `needs-reanchor` rather than guessing.
  - Show deck/slide/object-level AI activity without moving the user's live selection or exposing raw content in notifications.

### Acceptance criteria

- Manual editing works with MCP disabled, every action is undoable, and the canonical scene plus exported editable PPTX represent the same supported properties.
- A pin follows an unambiguous object through reflow; a deleted or ambiguously replaced target keeps its original reference crop and becomes `needs-reanchor`.
- A model can inspect one exact thread and current render, stage a bounded fix, and show a proposal on the same canvas without applying it.
- The submitted thread contains the exact slide revision, stable target ID or normalized region, reference crop, user comment, and permitted operation scope; the AI does not need to infer which object the user meant.
- A stale AI proposal cannot overwrite a newer human edit, and an object under active human manipulation cannot be changed by a proposal.

- [ ] **Build a first-class native table editor, layout solver, and export QA gate.**
  - Preserve stable table/row/column/cell IDs, exact text, merged topology, semantic color roles, native PowerPoint identity, captions/sources, and source content/structure hashes.
  - Add cell/range, row, and column selection; in-place editing; widths/heights; padding; horizontal/vertical alignment; merge/unmerge under policy; semantic roles; Template Pack styles; and approved-exemplar previews.
  - Implement Match approved exemplar, Fit to width, Balance columns, Distribute rows, Restore source structure, and continuation-slide proposals with repeated header semantics.
  - Measure every cell using the final font and width; validate wrap, clipping, minimum type, padding, row height, column balance, merge constraints, borders, contrast, semantic color, caption/source separation, and neighboring-object alignment.
  - Expose bounded structured table context and table crops to authorized MCP clients, and accept updates only against explicit stable IDs and the current revision.
  - Revalidate exact cell/structure hashes and optical quality after editable PPTX export; never hide/drop rows, rasterize a supported table, or silently shrink below the profile minimum.

### Acceptance criteria

- Synthetic merged, semantic-color, dense, and continuation tables remain native/editable and retain exact cells, order, and merge topology through project save/reopen and PPTX export.
- No exported cell is clipped, missing, below minimum type, or visually detached from its row/column/header role.
- Table findings link to the exact affected cell or edge and can be addressed by a human or a bounded AI proposal on the same canvas.
- An incompatible exemplar is rejected or role-adapted visibly; it is never forced onto a table in a way that changes meaning or structure.

- [ ] **Surface AI/MCP activity and pending approvals in the desktop UI.**
  - **Implemented foundation (2026-08-12):** active/completed/failed MCP operations now produce an accessible, color-independent desktop notification with a sanitized operation name and no raw content. The persistent approval queue, affected-deck links, dismissal controls, and state-matrix automation remain.
  - Show a restrained, color-independent activity indicator while an MCP tool is reading, auditing, or staging a proposal.
  - Show a completion toast when an MCP operation finishes, including the affected deck and whether anything changed.
  - Keep a persistent approval queue or badge for actions that require the user, such as confirming a target template, registering a style exemplar, and applying or rejecting a staged proposal.
  - Let the user open the exact deck and approval control directly from the notification.
  - Distinguish clearly between read-only inspection, a staged proposal, an applied change, and an exported file.
  - Do not display raw prompts, source excerpts, Resource contents, or sensitive presentation text in notifications or activity history.
  - Include accessible names, keyboard access, focus handling, and non-color status cues.

### Acceptance criteria

- A user can tell when an MCP client is actively using the open project.
- A user can identify every approval currently blocking an MCP-requested workflow without leaving the app to ask the AI client.
- Dismissing a toast does not dismiss or approve the underlying request; pending approvals remain visible until resolved.
- Automated tests cover active, completed, failed, stale, and approval-required MCP states.
