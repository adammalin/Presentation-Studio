# Presentation Studio TODO

## Authorized ORNL Template Pack

- [ ] **Install, version, and qualify the latest authorized ORNL PowerPoint template locally.**
  - **Implemented foundation (2026-08-11):** local POTX/PPTX selection, independent SHA-256 verification, protected app-data storage, restart persistence, full 30-layout inventory for the current ORNL reference, real template-media reuse, searchable/category-filtered gallery, selectable detail view, and app-only placeholder guides.
  - **Still required:** authoritative revision/effective-date/owner metadata, side-by-side pack selection and rollback, project-pinned pack snapshots, semantic slot mapping, compatibility scoring, staged reflow, and PowerPoint-native render comparison.
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
- Representative synthetic content is compiled and rendered through every discovered layout, with placeholder, theme, master, overflow, and editable-PowerPoint checks.
- Reopening a project uses its pinned Template Pack snapshot even after another authorized revision is installed.
- Repository and history scans find no official ORNL template, extracted template artwork, or production Template Pack bytes.

## Product roadmap

### Round goal and implementation order

The current implementation makes source and deterministic proposal designs visible before broad authoring features. Continue in this order:

1. ship the versioned Presentation Design Standard and one-click ORNL defaults;
2. make the Slides tab a real rendered gallery with a selectable close-up viewer;
3. apply the standard through the Designer Cleanup engine, beginning with native tables;
4. add anchored comments and **Submit to AI** on the same rendered slide surface;
5. add broader direct-manipulation editing after the shared scene, stable IDs, and comment anchors are reliable.

- [ ] **Ship one versioned Presentation Design Standard as the app, preflight, and MCP source of truth.**
  - **Implemented foundation (2026-08-12):** one shared versioned JSON standard now drives new-project defaults, resolved ORNL deck profiles, the Rules view, table fallback tokens, drift tests, and MCP cleanup/design-contract responses. Ordinary current/older ORNL imports adopt 16:9 + Aptos + exact-content cleanup without a target-template questionnaire. Schema negotiation, generated preflight/report views, and full cleanup-engine enforcement remain.
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
  - **Implemented foundation (2026-08-12):** the Slides tab renders every current embedded PPTX slide locally from OOXML, inherited layout geometry, editable text/shapes, images, connectors, and native table cells; thumbnails open a large revision-labeled review surface. Designer Cleanup reviews every slide, assigns change/approved-as-is/needs-review dispositions, repairs high-confidence cover text that drifts from a repeated peer edge, materializes selected changes in memory, and shows side-by-side Current/Proposal designs before acceptance. MCP can request either representation as a bounded JPEG and stage the full deck-wide proposal. Broader peer-group alignment, optical alignment, Source/Current semantic distinctions, Export comparison, zoom/filter context, layout remapping, full geometry reflow, and independent PowerPoint export rendering remain.
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
  - **Implemented foundation (2026-08-12):** Comment mode supports click/drag normalized-region anchors on the current slide, private saved notes, submitted AI threads, revision binding, project-package persistence, and bounded MCP list/read tools that exclude private notes. Direct manipulation, stable object/text/table-cell anchors, reference crops, replies, proposal application, resolution states, and undo/redo remain.
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
