import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { PresentationAppClient, PresentationAppUnavailableError } from "./presentation-app-client.mjs";
import { DESIGN_CONTRACT, designContractMessage } from "./design-contract.mjs";

const server = new McpServer({ name: "presentation-studio-local", version: "0.1.0" });
const client = new PresentationAppClient();

function success(result, message) {
  return { structuredContent: result, content: [{ type: "text", text: message }] };
}

function failure(error) {
  const message = error instanceof PresentationAppUnavailableError ? error.message : error instanceof Error ? error.message : "Presentation Studio could not complete the tool request.";
  return { isError: true, content: [{ type: "text", text: message.slice(0, 700) }] };
}

async function call(operation, input, message) {
  try {
    const result = await client.command(operation, input);
    return success(result, typeof message === "function" ? message(result) : message);
  } catch (error) { return failure(error); }
}

async function callImage(operation, input, message) {
  try {
    const result = await client.command(operation, input);
    const { data, mimeType, ...structuredContent } = result;
    if (typeof data !== "string" || !["image/jpeg", "image/png"].includes(mimeType)) throw new Error("Presentation Studio did not return a supported bounded slide image.");
    return { structuredContent, content: [{ type: "text", text: typeof message === "function" ? message(result) : message }, { type: "image", data, mimeType }] };
  } catch (error) { return failure(error); }
}

server.registerTool("get_design_contract", {
  title: "Get the presentation designer contract",
  description: "Read the mandatory deck-wide design and QA instructions that govern any AI model using Presentation Studio. Call this before proposing or performing cleanup. The contract requires improving every slide, preserving approved content exactly, inspecting every text box and visual, choosing the best approved layout, minimizing routine approval questions, and independently rendering the export for visual QA.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => success(DESIGN_CONTRACT, designContractMessage()));

server.registerTool("get_app_status", {
  title: "Get Presentation Studio status",
  description: "Check whether the local desktop app is open and summarize the active project without reading deck content. AI session access may remain off for this status check.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("get_app_status", {}, (result) => `Presentation Studio is open with ${result.project.deckCount} decks and ${result.project.slideCount} audited slides. AI session access is ${result.aiSessionAccess ? "on" : "off"}.`));

server.registerTool("list_decks", {
  title: "List presentation review jobs",
  description: "List deck IDs, names, operation scopes, template decisions, audit counts, status, and the exact updatedAt stale-write guard. Requires visible AI session access in Presentation Studio. Does not return slide text or source files.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("list_decks", {}, (result) => `Read ${result.decks.length} deck review job${result.decks.length === 1 ? "" : "s"}.`));

server.registerTool("get_deck_scene_summary", {
  title: "Read a deck's hybrid scene summary",
  description: "Read the versioned Presentation Studio scene revision, slide/object counts, fidelity-state totals, template binding, and PowerPoint preservation contract without returning slide text or package bytes. Call this before design work to learn what is editable-native, preserved-native, conversion-required, or blocked.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_scene_summary", input, (result) => `Read scene ${result.scene.revision} for ${result.deck.name}, containing ${result.scene.objectCount} source-bound objects.`));

server.registerTool("get_slide_scene", {
  title: "Read one slide's structured scene",
  description: "Read every source-bound object on one slide in z-order with semantic role, exact geometry, source locator, fidelity reason, represented properties, and currently permitted operations. This returns hashes but no visible text. Pair it with get_slide_design_context and the authoritative Current render; never edit a preserved or conversion-required internal structure as if it were native Studio content.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_slide_scene", input, (result) => `Read ${result.objects.length} source-bound objects on slide ${result.slide.number} of ${result.deck.name}.`));

server.registerTool("list_resources", {
  title: "List authorized project Resources",
  description: "List metadata for project Resources that a person explicitly shared for the current app session. Requires AI session access. Never returns original file bytes or extracted document text.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("list_resources", {}, (result) => `Read metadata for ${result.resources.length} of ${result.totalResourceCount} project Resource${result.totalResourceCount === 1 ? "" : "s"} authorized for this session.`));

server.registerTool("get_template_layout_catalog", {
  title: "Read the active Template Pack layout system",
  description: "Read every approved layout in the installed local PowerPoint Template Pack as semantic intent, editable slots, accepted content kinds, capacity, and fit constraints. This exposes no template bytes or artwork. Use it before choosing a layout; never select from the layout name alone.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("get_template_layout_catalog", {}, (result) => `Read ${result.layouts.length} semantic layouts from ${result.template.name}.`));

server.registerTool("get_template_layout_render", {
  title: "View one approved template layout",
  description: "Return the authoritative Microsoft PowerPoint render for one layout ID from the active local Template Pack, plus its semantic slots and renderer provenance. Read the catalog first. Empty image, table, chart, or flexible-content placeholders may appear blank until bound, so interpret the pixels together with the returned slot geometry.",
  inputSchema: { layoutId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_template_layout_render", input, (result) => `Rendered approved layout ${result.layout.ordinal}, ${result.layout.name}, at ${result.width} × ${result.height}.`));

server.registerTool("recommend_slide_layouts", {
  title: "Rank approved layouts for one slide",
  description: "Build an exact-content profile from one audited slide and deterministically rank the active Template Pack layouts by title, text-block, image, table, chart, media, density, and semantic-intent compatibility. Use the shortlist with Current and Proposal PowerPoint-native renders; a score is evidence, not final visual approval.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    slideNumber: z.number().int().min(1),
    limit: z.number().int().min(1).max(12).default(6),
  },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("recommend_slide_layouts", input, (result) => `Ranked ${result.recommendations.length} approved layouts for slide ${result.slide.number} of ${result.deck.name}.`));

server.registerTool("get_slide_design_work_order", {
  title: "Read one versioned AI slide-design work order",
  description: "Assemble the authoritative Current PowerPoint evidence, exact locked copy, hybrid scene objects and allowed operations, slide findings, submitted comments, ORNL rules, and ranked Template Pack layouts into one revision-bound work order. Call this before designing a slide; it is the primary Inspect and Diagnose input for the iterative visual-design loop.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_slide_design_work_order", input, (result) => `Built design work order ${result.revision} for slide ${result.slide.number} of ${result.deck.name}.`));

server.registerTool("get_deck_design_work_order", {
  title: "Read the representative deck-design qualification set",
  description: "Select up to five representative slides—cover, text-led content, diagram, image-heavy, and dense table—and return a complete versioned work order for each. Use this bounded set to prove the design loop before expanding across the full deck. Exact source content is included, so visible AI session access is required.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_design_work_order", input, (result) => `Built ${result.workOrders.length} representative design work orders for ${result.deck.name}.`));

server.registerTool("get_deck_audit", {
  title: "Read a deck audit",
  description: "Read deterministic template, font, editable text-box geometry, fit risk, alignment, object, comments, support, and production findings for one audited deck. Requires AI session access. Visible slide text and Resource bytes are intentionally omitted.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_audit", input, (result) => `Read the audit for ${result.deck.name} with ${result.audit.findings.length} findings.`));

server.registerTool("get_slide_design_context", {
  title: "Read bounded slide design context",
  description: "Read exact approved text mapped to stable editable object IDs, typography, geometry in inches, deterministic fit checks, safe alignment candidates, staged geometry, and object inventory for up to 10 consecutive slides. Use this with both Current and Proposal renders before staging an atomic multi-object layout so the model can make measured design decisions without rewriting content. Requires visible AI session access in Presentation Studio.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    startSlide: z.number().int().min(1),
    endSlide: z.number().int().min(1),
  },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_slide_design_context", input, (result) => `Read exact design context for slides ${result.range.start}–${result.range.end} of ${result.deck.name}.`));

server.registerTool("get_cleanup_rule_profile", {
  title: "Read the resolved cleanup rule profile",
  description: "Read the exact versioned Presentation Design Standard, one-click defaults, resolved deck profile, native-table tokens, autonomy policy, and template decision source. Requires AI session access and does not change the project.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_cleanup_rule_profile", input, (result) => `Read Presentation Design Standard ${result.standard.version} for ${result.deck.name}.`));

server.registerTool("get_pending_proposal_manifest", {
  title: "Read the pending proposal command manifest",
  description: "Return the current pending proposal's revision, selected command kinds, source-bound object IDs, native geometry, layout references, text-style tokens, vector-decoration geometry, and validation rationale without returning slide text or presentation bytes. Use it to audit a complex multi-slide transaction and diagnose materialization or native-render failures before acceptance.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_pending_proposal_manifest", input, (result) => `Read ${result.proposal.changes.length} selected and unselected command groups from pending proposal ${result.proposal.id}.`));

server.registerTool("get_slide_render", {
  title: "View a source, current, proposed, or export slide design",
  description: "Return a bounded revision-labeled image with explicit renderer provenance. Presentation Studio prefers an authoritative local Microsoft PowerPoint render and clearly labels the OOXML reconstruction when native rendering is unavailable. Read Source/Current and Proposal renders before staging or accepting visual changes; never treat an approximate fallback as final visual evidence.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), representation: z.enum(["source", "current", "proposal", "export"]).default("current") },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_slide_render", input, (result) => `Rendered ${result.representation} slide ${result.slide.number} of ${result.deck.name} at ${result.width} × ${result.height}.`));

server.registerTool("get_slide_render_comparison", {
  title: "Compare authoritative Current and Proposal renders",
  description: "Return one side-by-side Microsoft PowerPoint image plus raster hashes, changed-pixel ratio, mean channel delta, and changed-region bounds for a staged proposal. Pixel difference proves a visible change but does not prove the proposal is better; inspect both compositions and revise or reject regressions. Requires a current proposal and native PowerPoint rendering.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_slide_render_comparison", input, (result) => `Compared authoritative Current and Proposal renders for slide ${result.slide.number} of ${result.deck.name}; ${(result.metrics.changedPixelRatio * 100).toFixed(2)}% of pixels materially changed.`));

server.registerTool("reject_design_proposal", {
  title: "Reject an AI design draft after native review",
  description: "Reject only the currently pending proposal after inspecting its authoritative Microsoft PowerPoint Current/Proposal comparison. This records the AI's concrete visual-regression rationale and raster evidence, returns the deck to audited state, and leaves source bytes, accepted state, saved project files, and exports untouched. Use this during bounded autonomous iteration when a draft is not visibly better; it can never apply or export a proposal.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    proposalId: z.string().min(1).max(120),
    slideNumber: z.number().int().min(1),
    rationale: z.string().min(1).max(1_000),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("reject_design_proposal", input, (result) => `Rejected proposal ${result.proposal.id} after PowerPoint-native review. The source, accepted state, saved project, and exports are unchanged.`));

server.registerTool("list_design_threads", {
  title: "List design comments submitted to AI",
  description: "List only location-anchored design threads the user explicitly submitted to AI. Private notes are excluded. Results include the exact slide revision and normalized region but no unrelated slide content.",
  inputSchema: { deckId: z.string().min(1).max(120).optional() },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("list_design_threads", input, (result) => `Read ${result.threads.length} design thread${result.threads.length === 1 ? "" : "s"} submitted to AI.`));

server.registerTool("get_design_thread", {
  title: "Read one submitted design comment",
  description: "Read one exact submitted design thread, its revision-bound normalized region, current deck/slide locator, and bounded instruction. Read get_slide_render before proposing a fix; never guess if the anchor becomes ambiguous.",
  inputSchema: { threadId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_design_thread", input, (result) => `Read the submitted design thread on slide ${result.thread.slideNumber} of ${result.deck?.name ?? "the open deck"}.`));

server.registerTool("stage_font_cleanup", {
  title: "Stage conservative font cleanup",
  description: "Stage supported legacy-font mappings for visible human review in the app. The correct ORNL target must already be confirmed by a person. Pass the exact updatedAt returned by list_decks. This does not apply, save, or export changes.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().min(1).max(80) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_font_cleanup", input, (result) => `Staged ${result.proposal.summary} for human review. Nothing was applied, saved, or exported.`));

server.registerTool("stage_designer_cleanup", {
  title: "Stage deck-wide designer cleanup",
  description: "Review every slide and every directly editable text box. Stage supported exact-content typography, collision-checked high-confidence cover/peer text-box alignment, and native-table improvements; flag overflow risk, off-slide text, safe-margin cases, ambiguous alignment, complex tables, and semantic colors instead of hiding content or guessing. Record an explicit approved-as-is or needs-review disposition for every slide. This creates a reversible proposal with Current/Proposal renders; it does not apply, save, export, or overwrite anything.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_designer_cleanup", input, (result) => `Staged a deck-wide designer proposal covering ${result.proposal.slideDispositions.length} slides. Nothing was applied, saved, or exported.`));

server.registerTool("stage_table_design_update", {
  title: "Stage shared ORNL table components",
  description: "Apply one shared, versioned ORNL native-table component to 1–40 exact table IDs from get_slide_design_context. Use standard for ordinary tables and dense-technical for large technical tables that require smaller measured type and tighter padding. The component keeps exact cell text, order, merged structure, and editability while standardizing Aptos typography, header treatment, body fills, padding, and restrained rules across slides. Pass the exact IDs of submitted comments this change fully addresses in addressedThreadIds; those comments are removed from the clean canvas while unrelated comments remain active. This accumulates in the reversible proposal and never applies, saves, exports, or overwrites the source.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    tableIds: z.array(z.string().min(1).max(180)).min(1).max(40),
    variant: z.enum(["standard", "dense-technical"]),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_table_design_update", input, (result) => `Staged the shared ${result.variant} ORNL component for ${result.tableCount} native table${result.tableCount === 1 ? "" : "s"}. Compare Current and Proposal; nothing was applied, saved, exported, or overwritten.`));

server.registerTool("stage_slide_geometry_update", {
  title: "Stage a measured slide-object edit",
  description: "Stage a reversible move and/or resize for one stable editable object returned by get_slide_design_context. First inspect the Current render and object inventory, calculate exact geometry in inches, and explain the alignment or fit rationale. Pass only the exact submitted comment IDs this edit fully addresses in addressedThreadIds; those comments disappear from the clean slide while unrelated feedback stays active. The object must remain on-slide; protected slides, unsupported grouped resizing, stale revisions, and ambiguous IDs are rejected. This opens Current/Proposal review and never rewrites text, applies, saves, exports, or overwrites the source.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    objectId: z.string().min(1).max(180),
    xInches: z.number().min(0).max(20),
    yInches: z.number().min(0).max(20),
    widthInches: z.number().min(0.1).max(20),
    heightInches: z.number().min(0.1).max(20),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
    allowIntentionalOverlap: z.boolean().default(false),
    allowFitRisk: z.boolean().default(false),
    allowSafeArea: z.boolean().default(false),
    allowAspectRatioChange: z.boolean().default(false),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_slide_geometry_update", input, (result) => `Staged a measured ${result.object.kind} geometry proposal on slide ${result.object.slideNumber}. Compare Current and Proposal; nothing was applied, saved, exported, or overwritten.`));

server.registerTool("stage_slide_layout_update", {
  title: "Stage an atomic multi-object slide layout",
  description: "Stage up to 20 measured object moves/resizes or explicitly reset prior one-off geometry commands back to their native source positions for one slide as one atomic design transaction. Read get_slide_design_context plus the Current render first, preserve exact text, use the stable object IDs, and calculate all final coordinates together. Pass only the exact submitted comment IDs this layout fully addresses in addressedThreadIds; those comments disappear from the clean slide while unrelated feedback stays active. Presentation Studio rejects stale IDs, protected slides, off-slide geometry, unsafe picture distortion, worsened text fit, safe-margin regressions, and newly increased overlap unless the corresponding exception is explicitly authorized with a concrete rationale. The full layout appears in Current/Proposal review and remains unapplied, unsaved, and unexported.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
    resetObjectIds: z.array(z.string().min(1).max(180)).max(20).default([]),
    commands: z.array(z.object({
      objectId: z.string().min(1).max(180),
      xInches: z.number().min(0).max(20),
      yInches: z.number().min(0).max(20),
      widthInches: z.number().min(0.1).max(20),
      heightInches: z.number().min(0.1).max(20),
      rationale: z.string().min(1).max(700),
      allowIntentionalOverlap: z.boolean().default(false),
      allowFitRisk: z.boolean().default(false),
      allowSafeArea: z.boolean().default(false),
      allowAspectRatioChange: z.boolean().default(false),
    })).max(20).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_slide_layout_update", input, (result) => `Staged ${result.commands.length} validated object edits and ${result.resetObjectIds.length} source-geometry resets on slide ${result.slideNumber} as one proposal. Compare Current and Proposal; nothing was applied, saved, exported, or overwritten.`));

server.registerTool("stage_slide_visual_design", {
  title: "Stage native ORNL visual polish",
  description: "Stage editable PowerPoint-native text hierarchy plus restrained ORNL vector rules, panels, and frames for one slide. Use only source-bound text object IDs and non-text decorative shapes; this operation cannot add or rewrite claims. Inspect the authoritative Current render first, use the Presentation Design Standard and a fresh revision, keep all geometry on-slide, and prefer a few purposeful grouping cues over decoration. Pass only the exact submitted comment IDs this visual pass fully addresses in addressedThreadIds; those comments disappear from the clean slide while unrelated feedback stays active. The commands accumulate with other pending slide edits in one exact-content proposal and remain unapplied, unsaved, and unexported until reviewed.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    clearPendingLayoutRemap: z.boolean().default(false),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
    removeDecorationIds: z.array(z.string().min(1).max(120)).max(30).default([]),
    textStyles: z.array(z.object({
      objectId: z.string().min(1).max(180),
      fontSizePt: z.number().min(10).max(60).optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      color: z.string().regex(/^#[0-9A-F]{6}$/i).optional(),
      alignment: z.enum(["left", "center", "right"]).optional(),
      verticalAlignment: z.enum(["top", "middle", "bottom"]).optional(),
      insetsInches: z.object({ top: z.number().min(0).max(.25), right: z.number().min(0).max(.25), bottom: z.number().min(0).max(.25), left: z.number().min(0).max(.25) }).optional(),
      paragraphStyle: z.object({
        lineSpacingMultiple: z.number().min(.8).max(1.6).optional(),
        spaceAfterPt: z.number().min(0).max(30).optional(),
        bulletLeftMarginInches: z.number().min(0).max(1).optional(),
        bulletHangingInches: z.number().min(0).max(.5).optional(),
      }).optional(),
      rationale: z.string().min(1).max(700),
    })).max(20).default([]),
    decorations: z.array(z.object({
      id: z.string().min(1).max(120),
      name: z.string().min(1).max(120),
      xInches: z.number().min(0).max(20),
      yInches: z.number().min(0).max(20),
      widthInches: z.number().min(.01).max(20),
      heightInches: z.number().min(.01).max(20),
      fillColor: z.string().regex(/^#[0-9A-F]{6}$/i).optional(),
      lineColor: z.string().regex(/^#[0-9A-F]{6}$/i).optional(),
      lineWidthPt: z.number().min(0).max(6).default(0),
      behindContent: z.boolean().default(true),
      rationale: z.string().min(1).max(700),
    })).max(30).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_slide_visual_design", input, (result) => `Staged ${result.textStyleCount} text hierarchy updates, ${result.decorationCount} native vector design elements, and ${result.removedDecorationCount} decoration removals on slide ${result.slideNumber}${result.clearedLayoutRemap ? ", replacing the failed pending layout remap" : ""}. Compare Current and Proposal; nothing was applied, saved, exported, or overwritten.`));

server.registerTool("stage_slide_native_layout", {
  title: "Stage a real approved PowerPoint layout",
  description: "Stage one Designer Cleanup reflow to a real approved Template Pack layout. Presentation Studio first reuses an exact SHA-256-matching native layout already present in the deck; otherwise it clones the guarded master/layout/theme/media dependency graph. It remaps compatible placeholder identities and repoints only the named slide. Exact slide text, native objects, source package, and all other slide relationships remain protected. Pass only the exact submitted comment IDs this remap fully addresses in addressedThreadIds; those comments disappear from the clean slide while unrelated feedback stays active. Requires a fresh design work order, compatible layout, active exact Template Pack revision, and reflow scope. The result is rendered through Microsoft PowerPoint for Current/Proposal review and must be rejected if it is visually worse; this tool never applies, saves, exports, or overwrites the source.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    workOrderRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    layoutId: z.string().min(1).max(120),
    rationale: z.string().min(1).max(1_000),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
    allowPoorLayout: z.boolean().default(false),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_slide_native_layout", input, (result) => `Staged approved native layout ${result.targetLayout.name} on slide ${result.slideNumber}. Inspect the PowerPoint-native comparison; nothing was applied, saved, exported, or overwritten.`));

server.registerTool("stage_slide_recomposition", {
  title: "Stage semantic recomposition into an approved layout",
  description: "Atomically apply one real approved Template Pack layout and bind 1 to 20 compatible source-bound objects to its unique semantic slots. Presentation Studio reuses an exact SHA-256-matching native layout already in the deck or clones the guarded dependency graph, remaps compatible placeholders, computes native PowerPoint geometry, preserves exact content and unbound objects, rejects stale work orders, incompatible slots, crop-unsafe pictures, substantial unmeasured text-frame replacement, overlap, worsened text fit, and unsupported operations, then opens Current/Proposal review. Pass only the exact submitted comment IDs this recomposition fully addresses in addressedThreadIds; those comments disappear from the clean slide while unrelated feedback stays active. Prefer align-horizontal for existing text frames: it adopts the approved layout's horizontal edges while preserving proven vertical fit. Use fill only for bounded text frames that do not substantially shrink or move vertically, and contain for pictures. Reject any visually weaker PowerPoint-native proposal.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    workOrderRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    layoutId: z.string().min(1).max(120),
    rationale: z.string().min(1).max(1_000),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
    allowPoorLayout: z.boolean().default(false),
    bindings: z.array(z.object({
      objectId: z.string().min(1).max(180),
      slotId: z.string().min(1).max(180),
      fit: z.enum(["fill", "contain", "align-horizontal"]).default("align-horizontal"),
      insetInches: z.number().min(0).max(.25).default(0),
    })).min(1).max(20),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_slide_recomposition", input, (result) => `Staged ${result.bindings.length} semantic bindings for slide ${result.slideNumber} using ${result.targetLayout.name}. Inspect the native comparison; nothing was applied, saved, exported, or overwritten.`));

await server.connect(new StdioServerTransport());
