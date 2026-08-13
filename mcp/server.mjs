import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { PresentationAppClient, PresentationAppUnavailableError } from "./presentation-app-client.mjs";
import { DESIGN_CONTRACT, designContractMessage } from "./design-contract.mjs";
import { stripImagePayloads } from "./image-payload.mjs";

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

async function callImages(operation, input, message) {
  try {
    const result = await client.command(operation, input);
    const images = Array.isArray(result.images) ? result.images : [];
    if (!images.length || images.some((item) => typeof item?.data !== "string" || !["image/jpeg", "image/png"].includes(item?.mimeType))) throw new Error("Presentation Studio did not return the required bounded inspection images.");
    const structuredContent = { ...result, images: stripImagePayloads(images) };
    return {
      structuredContent,
      content: [
        { type: "text", text: typeof message === "function" ? message(result) : message },
        ...images.map((item) => ({ type: "image", data: item.data, mimeType: item.mimeType })),
      ],
    };
  } catch (error) { return failure(error); }
}

server.registerTool("get_design_contract", {
  title: "Get the presentation designer contract",
  description: "Read the mandatory web-first presentation design and QA instructions that govern any AI model using Presentation Studio. Call this before design work. The contract requires improving every slide by importing exact content into the shared Studio HTML/CSS scene, making a substantive whole-slide composition decision with shared ORNL components or an approved layout, preserving approved content exactly, minimizing routine approval questions, compiling to editable PowerPoint, and independently rendering the result for visual QA.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => success(DESIGN_CONTRACT, designContractMessage()));

server.registerTool("get_app_status", {
  title: "Get Presentation Studio status",
  description: "Check whether the local desktop app is open, summarize the active project without reading deck content, and report whether PowerPoint-native render/measurement is ready or blocked by a locked Mac. AI session access may remain off for this status check.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("get_app_status", {}, (result) => `Presentation Studio is open with ${result.project.deckCount} decks and ${result.project.slideCount} audited slides. AI session access is ${result.aiSessionAccess ? "on" : "off"}. PowerPoint-native QA is ${result.nativePowerPoint?.ready ? "ready" : result.nativePowerPoint?.sessionLocked ? "blocked until the Mac is unlocked" : "unavailable"}.`));

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

server.registerTool("get_studio_web_scene", {
  title: "Read one slide as a Studio web-design scene",
  description: "Read the canonical semantic HTML/CSS scene for one imported PowerPoint slide: exact locked text and table cells, source-bound editable node IDs, source and current frames in inches, semantic roles, media bindings, ORNL design tokens, and the recommended shared layout recipe. This is the AI and human design surface—use it to make an actual composition decision instead of merely shrinking fonts or preserving weak source coordinates. The source PowerPoint remains the preservation envelope and its native render remains final visual authority.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_studio_web_scene", input, (result) => `Read ${result.slide.nodes.length} semantic web nodes for slide ${result.slide.slideNumber} of ${result.deck.name}; recommended recipe: ${result.slide.recommendedRecipe}.`));

server.registerTool("list_resources", {
  title: "List authorized project Resources",
  description: "List metadata for project Resources that a person explicitly shared for the current app session. Requires AI session access. Never returns original file bytes or extracted document text.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("list_resources", {}, (result) => `Read metadata for ${result.resources.length} of ${result.totalResourceCount} project Resource${result.totalResourceCount === 1 ? "" : "s"} authorized for this session.`));

server.registerTool("get_template_layout_catalog", {
  title: "Read the active Template Pack layout system",
  description: "Read every approved layout in the installed local PowerPoint Template Pack as semantic intent plus responsive editable-slot contracts: preferred/minimum/maximum bounds, allowed object kinds, alignment and padding intent, priority, capacity, and fit constraints. This exposes no template bytes or artwork. Use it before choosing a layout; never select from the layout name alone.",
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
}, (input) => callImage("get_slide_design_work_order", input, (result) => `Built image-bearing design work order ${result.revision} for slide ${result.slide.number} of ${result.deck.name}. The attached PNG is an authoritative 2,200-pixel Microsoft PowerPoint render.`));

server.registerTool("get_slide_inspection_packet", {
  title: "Inspect a slide with native pixels, crops, measurements, and metrics",
  description: "Return one revision-bound inspection packet containing the exact design work order, a 2,200-pixel PowerPoint-native full-slide PNG, readable title/table/text crops, a deterministic crop overlay, native rendered-text and cell measurements, and objective design metrics. Call this instead of guessing geometry from a screenshot. Use current before design work and proposal after staging; pixels guide gestalt, PowerPoint supplies measurements, and deterministic solvers supply exact coordinates.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), representation: z.enum(["current", "proposal"]).default("current") },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImages("get_slide_inspection_packet", input, (result) => `Built ${result.representation} inspection packet ${result.revision} with ${result.images.length} PowerPoint-native visual evidence images for slide ${result.slide.number} of ${result.deck.name}.`));

server.registerTool("get_deck_design_work_order", {
  title: "Read the representative deck-design qualification set",
  description: "Select up to five representative slides—cover, text-led content, diagram, image-heavy, and dense table—and return a complete versioned work order for each plus the deck-wide semantic table-color map. Use this bounded set to prove the design loop before expanding across the full deck. Exact source content is included, so visible AI session access is required.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_design_work_order", input, (result) => `Built ${result.workOrders.length} representative design work orders for ${result.deck.name}.`));

server.registerTool("get_deck_contact_sheet", {
  title: "Review a PowerPoint-native deck contact sheet",
  description: "Return one revision-bound page of up to 40 PowerPoint-native slide thumbnails for deck-level visual review. Page through the complete Current or Proposal deck to judge hierarchy, density, pacing, repeated-component consistency, and outliers before opening precise slide inspection packets. The contact sheet is authoritative for gestalt but not point geometry.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    representation: z.enum(["current", "proposal"]).default("current"),
    page: z.number().int().min(1).max(25).default(1),
  },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImages("get_deck_contact_sheet", input, (result) => `Rendered PowerPoint-native ${result.representation} deck contact sheet page ${result.page} of ${result.pageCount}, covering slides ${result.firstSlideNumber}–${result.lastSlideNumber} of ${result.deck.name}.`));

server.registerTool("get_deck_audit", {
  title: "Read a deck audit",
  description: "Read deterministic template, font, editable text-box geometry, fit risk, alignment, object, comments, support, and production findings for one audited deck. Requires AI session access. Visible slide text and Resource bytes are intentionally omitted.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_audit", input, (result) => `Read the audit for ${result.deck.name} with ${result.audit.findings.length} findings.`));

server.registerTool("get_slide_measurements", {
  title: "Read native PowerPoint slide measurements",
  description: "Read revision-bound Microsoft PowerPoint measurements for every source-bound object on one current or proposed slide: rendered text bounds, text-frame margins, line counts, object geometry, table row heights, column widths, cell bounds, cell margins, rendered cell-text bounds, and clearance metrics. The result returns hashes and counts but no slide copy. Use these facts and the semantic solvers instead of estimating point geometry from pixels.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), representation: z.enum(["current", "proposal"]).default("current") },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_slide_measurements", input, (result) => `Read ${result.measurement.authority} measurements and design metrics for ${result.representation} slide ${result.slideNumber} of ${result.deck.name}.`));

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

server.registerTool("record_proposal_visual_critique", {
  title: "Record a revision-bound AI visual critique",
  description: "After reading the Proposal inspection packet, record whether the native PowerPoint draft is visually better, needs another semantic revision, or should be rejected. The exact inspection revision, PowerPoint raster hashes, objective metric changes, rationale, and attempt number are persisted. A requested better verdict is withheld when deterministic metrics regress or pixels are unchanged. Automatic AI revision is capped at three attempts; attempt three rejects an unresolved draft. This never applies, saves, exports, or overwrites content.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    proposalId: z.string().min(1).max(120),
    slideNumber: z.number().int().min(1),
    inspectionRevision: z.string().min(1).max(500),
    verdict: z.enum(["better", "revise", "reject"]),
    rationale: z.string().min(1).max(1_000),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("record_proposal_visual_critique", input, (result) => `Recorded AI visual iteration ${result.critique.attempt}/3 as ${result.recordedVerdict}. The proposal remains ${result.proposal.status}; nothing was applied, saved, exported, or overwritten.`));

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

server.registerTool("stage_studio_web_design", {
  title: "Recompose a slide in Studio and stage editable PowerPoint",
  description: "Choose one shared Studio HTML/CSS layout recipe—or a real installed Template Pack layout—then atomically recompose the slide's exact source-bound content and compile its web-computed frames, typography, and ORNL components back to editable native PowerPoint objects. Optional nodeFrames are deliberate final refinements in inches, not a substitute for choosing a coherent recipe. Use get_studio_web_scene first. Make a substantive layout decision when the source is weak; do not default to smaller text or no-op cleanup. Submitted comments listed in addressedThreadIds are removed only when this proposal actually affects their slide. The result is a reversible Current/Proposal draft and never applies, saves, exports, or overwrites the source.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    recipe: z.enum(["source", "ornl-title-content", "ornl-title-two-column", "ornl-title-table", "ornl-title-figure-grid", "template-layout"]),
    layoutId: z.string().min(1).max(120).optional(),
    rationale: z.string().min(1).max(1000),
    nodeFrames: z.array(z.object({
      nodeId: z.string().min(1).max(180),
      xInches: z.number().min(0).max(20),
      yInches: z.number().min(0).max(20),
      widthInches: z.number().min(.1).max(20),
      heightInches: z.number().min(.1).max(20),
      rotation: z.number().min(-360).max(360).default(0),
    })).max(30).default([]),
    nodeStyles: z.array(z.object({
      nodeId: z.string().min(1).max(180),
      fontSizePt: z.number().min(10).max(60).optional(),
      fontWeight: z.union([z.literal(400), z.literal(600), z.literal(700)]).optional(),
      color: z.string().regex(/^#[0-9a-f]{6}$/i).optional(),
      textAlign: z.enum(["left", "center", "right"]).optional(),
      verticalAlign: z.enum(["top", "middle", "bottom"]).optional(),
      objectFit: z.enum(["contain", "cover"]).optional(),
    })).max(30).default([]),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_studio_web_design", input, (result) => `Recomposed slide ${result.slide.number} with ${result.slide.recipe} and staged ${result.proposal.geometryCount} source-bound PowerPoint geometry edits for native Current/Proposal review. Nothing was applied, saved, exported, or overwritten.`));

server.registerTool("stage_font_cleanup", {
  title: "Stage conservative font cleanup",
  description: "Stage supported legacy-font mappings for visible human review in the app. The correct ORNL target must already be confirmed by a person. Pass the exact updatedAt returned by list_decks. This does not apply, save, or export changes.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().min(1).max(80) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_font_cleanup", input, (result) => `Staged ${result.proposal.summary} for human review. Nothing was applied, saved, or exported.`));

server.registerTool("stage_designer_cleanup", {
  title: "Stage deck-wide designer cleanup",
  description: "Review every slide and every directly editable text box. Stage supported exact-content typography, collision-checked high-confidence cover/peer text-box alignment, and native-table improvements; flag overflow risk, off-slide text, safe-margin cases, ambiguous alignment, complex tables, and semantic colors instead of hiding content or guessing. Record an explicit approved-as-is or needs-review disposition for every slide. This creates a reversible proposal with Current/Proposal renders; it does not apply, save, export, or overwrite anything.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }), designStandardVersion: z.string().min(1).max(80) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_designer_cleanup", input, (result) => `Staged a deck-wide designer proposal covering ${result.proposal.slideDispositions.length} slides. Nothing was applied, saved, or exported.`));

server.registerTool("stage_table_design_update", {
  title: "Stage shared ORNL table components",
  description: "Apply one shared, versioned ORNL native-table component to 1–40 exact table IDs from get_slide_design_context. Echo the current designStandardVersion and use semanticColorPolicy preserve-source. Use standard for ordinary tables and dense-technical for large technical tables that require smaller measured type and tighter padding. The component keeps exact cell text, order, merged structure, editability, and every detected semantic role-to-cell mapping while adapting source accent fills only to their approved ORNL tints. It rejects a missing, stale, neutralized, or swapped semantic-color mapping. Pass the exact IDs of submitted comments this change fully addresses in addressedThreadIds; those comments are removed from the clean canvas while unrelated comments remain active. This accumulates in the reversible proposal and never applies, saves, exports, or overwrites the source.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    tableIds: z.array(z.string().min(1).max(180)).min(1).max(40),
    variant: z.enum(["standard", "dense-technical"]),
    designStandardVersion: z.string().min(1).max(80),
    semanticColorPolicy: z.literal("preserve-source"),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_table_design_update", input, (result) => `Staged the shared ${result.variant} ORNL component for ${result.tableCount} native table${result.tableCount === 1 ? "" : "s"}. Compare Current and Proposal; nothing was applied, saved, exported, or overwritten.`));

server.registerTool("solve_and_stage_alignment", {
  title: "Solve and stage semantic alignment",
  description: "Align 2–20 source-bound objects with a deterministic minimum-movement solver. Optical-left uses PowerPoint-native rendered text starts; structural modes use native object edges or centers. The solver enforces slide bounds, safe margins, and collision constraints, returns infeasible instead of guessing, and stages only a reversible Current/Proposal transaction. Use this for titles, body columns, captions, and peer elements instead of manually inventing x/y coordinates.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    objectIds: z.array(z.string().min(1).max(180)).min(2).max(20),
    anchorObjectId: z.string().min(1).max(180).optional(),
    mode: z.enum(["left", "optical-left", "center", "right", "top", "middle", "bottom"]),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("solve_and_stage_alignment", input, (result) => result.staged ? `Solved and staged ${result.result.commands.length} minimum-movement alignment edits. Inspect and remeasure the Proposal; nothing was applied, saved, exported, or overwritten.` : `The alignment solver refused the requested transaction: ${result.result.diagnostics.join(" ")}`));

server.registerTool("solve_and_stage_distribution", {
  title: "Solve and stage equal-gap distribution",
  description: "Distribute 3–20 source-bound objects—or three or more declared object groups—with exact horizontal or vertical equal gaps while preserving the outer span and minimizing movement. Use groups for cards, panels, captions, or other components whose children must move together. The deterministic solver rejects overlap, safe-region, and collision failures instead of asking the AI to estimate coordinates. The result remains a reversible Current/Proposal transaction.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    objectIds: z.array(z.string().min(1).max(180)).min(3).max(20),
    groups: z.array(z.array(z.string().min(1).max(180)).min(1).max(10)).min(3).max(10).optional(),
    mode: z.enum(["horizontal-equal-gap", "vertical-equal-gap"]),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("solve_and_stage_distribution", input, (result) => result.staged ? `Solved and staged ${result.result.commands.length} equal-gap distribution edits. Inspect and remeasure the Proposal; nothing was applied, saved, exported, or overwritten.` : `The distribution solver refused the requested transaction: ${result.result.diagnostics.join(" ")}`));

server.registerTool("solve_and_stage_safe_region", {
  title: "Fit objects into the slide safe region",
  description: "Move 1–20 source-bound objects as one intact group by the minimum distance needed to fit inside the 18-point slide safe region. The deterministic solver preserves every relative position, rejects infeasible oversized groups and new collisions, then remeasures the staged proposal in PowerPoint. Use this instead of inventing x/y coordinates for near-edge content.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    objectIds: z.array(z.string().min(1).max(180)).min(1).max(20),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("solve_and_stage_safe_region", input, (result) => result.staged ? `Solved and staged ${result.result.commands.length} minimum-movement safe-region edits. Inspect and remeasure the Proposal; nothing was applied, saved, exported, or overwritten.` : `The safe-region solver refused the requested transaction: ${result.result.diagnostics.join(" ")}`));

server.registerTool("solve_and_stage_group_layout", {
  title: "Compose visual groups in an approved layout region",
  description: "Place 1–10 declared visual groups into one responsive semantic slot from the active approved Template Pack as a horizontal or vertical stack. Presentation Studio preserves each group's internal geometry and relationships, applies versioned hierarchy-aware spacing for primary, supporting, and caption groups, honors slot padding and cross-axis alignment, and centers unused whitespace. When explicitly enabled, bounded proportional scaling can fit resize-capable groups down to a declared floor; PowerPoint remeasurement still rejects text-fit regression. Infeasible regions or new collisions are withheld. Use this after the AI chooses the appropriate layout, grouping, and semantic hierarchy; do not calculate coordinates manually.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    layoutId: z.string().min(1).max(180),
    slotId: z.string().min(1).max(180),
    groups: z.array(z.array(z.string().min(1).max(180)).min(1).max(10)).min(1).max(10),
    groupRoles: z.array(z.enum(["primary", "supporting", "caption"])).min(1).max(10).optional(),
    mode: z.enum(["horizontal-stack", "vertical-stack"]),
    alignment: z.enum(["start", "center", "end"]).default("start"),
    preferredGapPt: z.number().min(0).max(72).default(18),
    allowResponsiveScale: z.boolean().default(false),
    minimumScale: z.number().min(.5).max(1).default(.75),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("solve_and_stage_group_layout", input, (result) => result.staged ? `Solved and staged ${result.result.commands.length} approved-region group-layout edits. Inspect the native Proposal pixels; nothing was applied, saved, exported, or overwritten.` : `The approved-region group solver refused the requested transaction: ${result.result.diagnostics.join(" ")}`));

server.registerTool("fit_scene_to_layout", {
  title: "Fit a complete slide scene to approved layout regions",
  description: "Atomically fit 1–30 source-bound objects across 1–8 responsive semantic regions from one approved Template Pack layout. Declare relationship-preserving groups and primary/supporting/caption hierarchy per region; Presentation Studio applies each slot's preferred bounds, padding and alignment intent, optionally performs bounded proportional scaling on resize-capable groups, validates cross-region collisions and the slide safe area as one transaction, then rerenders and remeasures the reversible Proposal in PowerPoint. Use this shared-layout operation instead of redrawing each slide or calculating coordinates independently.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    layoutId: z.string().min(1).max(180),
    regions: z.array(z.object({
      slotId: z.string().min(1).max(180),
      groups: z.array(z.array(z.string().min(1).max(180)).min(1).max(10)).min(1).max(10),
      groupRoles: z.array(z.enum(["primary", "supporting", "caption"])).min(1).max(10).optional(),
      mode: z.enum(["horizontal-stack", "vertical-stack"]),
      alignment: z.enum(["start", "center", "end"]).optional(),
      preferredGapPt: z.number().min(0).max(72).optional(),
      allowResponsiveScale: z.boolean().default(false),
      minimumScale: z.number().min(.5).max(1).default(.75),
    })).min(1).max(8),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("fit_scene_to_layout", input, (result) => result.staged ? `Fitted and staged ${result.result.commands.length} cross-region scene-layout edits. Inspect the native Proposal pixels; nothing was applied, saved, exported, or overwritten.` : `The responsive scene-layout solver refused the requested transaction: ${result.result.diagnostics.join(" ")}`));

server.registerTool("solve_and_stage_table_layout", {
  title: "Solve and stage a native table layout",
  description: "Fit one cell-level native PowerPoint table using measured text bounds, row heights, column widths, cell margins, readable type floors, and minimum padding. The deterministic solver preserves a table that already passes, reallocates rows and columns when needed, and may grow the native table by the minimum amount that fits inside the safe region before PowerPoint remeasurement. It preserves exact cell copy and merged structure and returns concrete space recommendations when constraints remain infeasible instead of shrinking text. The solved grid remains a reversible Current/Proposal transaction and must be remeasured before acceptance.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    tableId: z.string().min(1).max(180),
    variant: z.enum(["standard", "dense-technical"]).default("standard"),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("solve_and_stage_table_layout", input, (result) => result.staged ? `Solved and staged a cell-level native table layout for ${result.tableId}. Inspect and remeasure the Proposal; nothing was applied, saved, exported, or overwritten.` : result.result.status === "already-fit" ? `PowerPoint confirms that ${result.tableId} already satisfies the resolved table constraints; no proposal was needed.` : `The table solver refused to violate its readability constraints. ${result.result.diagnostics.reasons.join(" ")} ${result.result.diagnostics.recommendations.join(" ")}`));

server.registerTool("solve_and_stage_text_fit", {
  title: "Fit text from native PowerPoint measurements",
  description: "Fit one source-bound text frame using PowerPoint-native rendered bounds and real frame margins. The deterministic solver raises uniform body or caption type to the resolved readability floor when needed, grows the frame by the minimum measured amount around its existing vertical anchor, preserves exact copy, and remeasures/rerenders the proposal in PowerPoint. It refuses horizontal clipping, mixed run-level hierarchy below the floor, safe-region regressions, and unresolved overflow; use an approved wider/taller region or semantic recomposition when infeasible. This remains a reversible Current/Proposal transaction and never silently shrinks text.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    objectId: z.string().min(1).max(180),
    rationale: z.string().min(1).max(700),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("solve_and_stage_text_fit", input, (result) => result.staged ? `PowerPoint remeasurement confirmed a non-clipping text fit for ${result.result.geometry?.objectId ?? result.result.textStyle?.objectId}. Inspect the Proposal pixels; nothing was applied, saved, exported, or overwritten.` : `The text-fit solver withheld the draft. ${result.result.diagnostics.reasons.join(" ")} ${result.result.diagnostics.recommendations.join(" ")}`));

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
