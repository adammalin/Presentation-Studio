import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { PresentationAppClient, PresentationAppUnavailableError } from "./presentation-app-client.mjs";
import { DESIGN_CONTRACT, designContractMessage } from "./design-contract.mjs";
import { stripImagePayloads } from "./image-payload.mjs";

const server = new McpServer({ name: "presentation-studio-local", version: "0.3.1" });
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
  description: "Read the mandatory web-first presentation design and QA instructions that govern any AI model using Presentation Studio. Call this before design work. The contract covers both improving every slide in an imported deck and creating a brand-new source-grounded deck directly from Resources in native Studio JSON. It requires a substantive whole-slide composition decision with shared ORNL components or an approved layout, preserved or source-grounded content, minimal routine approval questions, editable PowerPoint compilation, and independent visual QA.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => success(DESIGN_CONTRACT, designContractMessage()));

server.registerTool("get_agent_runbook", {
  title: "Get the current Presentation Studio agent runbook",
  description: "Return the concise current-state workflow for one open deck: non-negotiable ORNL and source-preservation rules, representative communication-archetype qualification set, explicit preserve/polish/recompose/rebuild-figure decisions, deck-consistency status, build and native-review readiness, and exactly one recommended next MCP action. Call this after get_design_contract and list_decks, and again whenever the central scene revision changes. It does not return presentation bytes, change a slide, save, or export.",
  inputSchema: { deckId: z.string().min(1).max(120).optional() },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_agent_runbook", input, (result) => `Read the current runbook for ${result.deck.name}. Next action: ${result.nextAction.tool}.`));

server.registerTool("get_app_status", {
  title: "Get Presentation Studio status",
  description: "Check whether the local desktop app is open, summarize the active project without reading deck content, and report whether PowerPoint-native render/measurement is ready or blocked by a locked Mac. AI session access may remain off for this status check.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("get_app_status", {}, (result) => `Presentation Studio is open with ${result.project.deckCount} decks and ${result.project.slideCount} audited slides. AI session access is ${result.aiSessionAccess ? "on" : "off"}. PowerPoint-native QA is ${result.nativePowerPoint?.ready ? "ready" : result.nativePowerPoint?.sessionLocked ? "blocked until the Mac is unlocked" : "unavailable"}.`));

server.registerTool("list_decks", {
  title: "List presentation review jobs",
  description: "List deck IDs, names, operation scopes, detected source-template classification, target template, decision source, resolved templateWorkflow, audit counts, status, and the exact updatedAt stale-write guard. ORNL is the automatic target unless the user explicitly selected preserve-source. Always obey templateWorkflow before choosing tools: source-template-cleanup preserves sponsor/custom masters through source-bound proposals; ornl-studio requires ORNL recipes/fresh composition; template-decision-required must be resolved first. Requires visible AI session access and returns no slide text or source files.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("list_decks", {}, (result) => `Read ${result.decks.length} deck review job${result.decks.length === 1 ? "" : "s"}.`));

server.registerTool("set_deck_template_target", {
  title: "Set a deck's design target",
  description: "Explicitly set the current deck to the default ORNL Studio workflow or preserve its detected source template. ORNL is the product default and should be selected without asking unless the user explicitly requests another brand or source-template preservation. Changing the target clears stale proposals and Studio compositions but never changes source bytes, saves, or exports.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime(),
    target: z.enum(["ornl-default", "preserve-source"]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("set_deck_template_target", input, (result) => `${result.deck.name} now targets ${result.deck.targetTemplateId}. Source PowerPoint bytes remain unchanged.`));

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
  description: "Read the canonical semantic HTML/CSS scene (an inspection scene for source-template-cleanup) at 13.333 × 7.5 inches for one imported PowerPoint slide: complete authored content excluding inherited slide-number/date/header/footer furniture, exact locked node text and table cells, semantic atom candidates, content-mapping coverage, binding provenance, source/current frames, repeated-component roles, media, and source-locked figure treatments. Obey list_decks.templateWorkflow: use this as the ORNL composition surface only for ornl-studio and make an actual composition decision there; for source-template-cleanup it is an inspection surface and sponsor/custom masters must be preserved through source-bound proposals. HARD RULE: an existing populated ORNL title slide is sacred and remains source-preserved exactly. If coverage is incomplete, fresh composition is held instead of silently omitting content. PowerPoint-native pixels remain final visual authority.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_studio_web_scene", input, (result) => `Read ${result.slide.nodes.length} semantic web nodes for slide ${result.slide.slideNumber} of ${result.deck.name}; recommended recipe: ${result.slide.recommendedRecipe}.`));

server.registerTool("get_studio_deck_consistency", {
  title: "Review repeated design systems across the Studio deck",
  description: "Read a deterministic revision-bound review of title-grid outliers, repeated component typography, related table structural styles, and unexplained one-off recipe/intervention patterns within the same communication archetype across the one canonical Studio scene. Findings identify exact slide and node IDs but make no changes. Use this after representative archetype qualification and before whole-deck qualification; preserve intentional semantic table-color differences and judge final appearance through PowerPoint-native pixels.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_studio_deck_consistency", input, (result) => `Reviewed ${result.review.designedSlideCount} designed slides in ${result.deck.name} and found ${result.review.issueCount} deck-system difference${result.review.issueCount === 1 ? "" : "s"}.`));

server.registerTool("publish_studio_component_style", {
  title: "Publish one reusable Studio component style",
  description: "Adopt the complete style of one source-bound repeated Studio component instance and propagate it to compatible instances with the same semantic role and light/dark surface class. This is the safe deck-system operation for repeated headings, captions, labels, objective copy, process stages, and related recipe components. It preserves every node's exact wording, data, geometry, source binding, semantic table colors, and protected ORNL template slides. Build affected slides and run get_studio_deck_consistency afterward; nothing is saved or exported.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    nodeId: z.string().min(1).max(180),
    name: z.string().min(1).max(120).optional(),
    targetSlideNumbers: z.array(z.number().int().min(1)).min(1).max(200).optional(),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("publish_studio_component_style", input, (result) => `Published ${result.component.name} to ${result.component.affectedNodeCount} compatible component instance${result.component.affectedNodeCount === 1 ? "" : "s"} across ${result.component.affectedSlideNumbers.length} slide${result.component.affectedSlideNumbers.length === 1 ? "" : "s"}. Exact content and geometry remain locked; build and inspect the affected PowerPoint results.`));

server.registerTool("preview_studio_fresh_composition", {
  title: "Build and view a fresh editable Studio composition",
  description: "For an ornl-studio target only, compile one already-designed Studio HTML/CSS slide into one or more genuinely new editable native PowerPoint output slides, validate exact source/output text using the authored-content inventory plus native table structure, render every written artifact slide with Microsoft PowerPoint, and return authoritative PNGs for critique. This deliberately replaces source coordinates and furniture, does not preserve the imported master, and therefore must never be used for source-template-cleanup unless the user explicitly selected ORNL cross-template conversion. It never applies, saves, exports, or overwrites the project or source file.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => callImages("preview_studio_fresh_composition", input, (result) => `Built ${result.editablePowerPoint.slideCount} fresh editable PowerPoint output slide${result.editablePowerPoint.slideCount === 1 ? "" : "s"} from source slide ${result.slide.number} of ${result.deck.name}, passed explicit source/output exact-copy guards, and rendered every written artifact slide authoritatively in Microsoft PowerPoint. Nothing was applied, saved, exported, or overwritten.`));

server.registerTool("get_studio_slide_critique", {
  title: "Find issues in the exact Studio export result",
  description: "Return the original PowerPoint slide and the exact current Studio export result as authoritative images, plus any explicitly preview-authorized concept-only references and deterministic PowerPoint-native overflow, safe-region, optical-alignment, spacing, hierarchy, density, and figure checks. Use this after building the exact slide revision. Compare message intent, labels, values, arrows, grouping, and causality to the original; compare only approved visual influences to a concept; then either refine with high-level Studio operations or record a bounded visual critique. This is the Found issues step and never changes, saves, or exports content.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), expectedSceneRevision: z.string().min(1).max(500) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImages("get_studio_slide_critique", input, (result) => `Found ${result.critique.issues.length} objective issue${result.critique.issues.length === 1 ? "" : "s"} on Studio slide ${result.slideNumber}; compare the original and export-result images before recording pass ${result.critique.iteration.currentPass}/3.`));

server.registerTool("repair_studio_objective_issues", {
  title: "Fix bounded PowerPoint-native Studio issues",
  description: "Apply one conservative deterministic Fixing pass to the exact issues returned by get_studio_slide_critique. Studio may minimally fit a complete relationship group into the safe region, replay a recorded optical alignment or equal-gap constraint, grow one editable text frame by the minimum PowerPoint-measured amount when collision-free, or restore title hierarchy within the ORNL type ceiling. It defers horizontal clipping, ambiguous figures, dense content, unsafe growth, table semantics, and material composition choices instead of shrinking everything or guessing coordinates. Exact wording, data, table structure, source-significant colors, and the sacred ORNL title slide remain locked. Every material change invalidates the old raster and requires a fresh PowerPoint build and critique; nothing is saved or exported.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    expectedRasterSha256: z.string().length(64),
    slideNumber: z.number().int().min(1),
    issueIds: z.array(z.string().min(1).max(180)).max(30).optional(),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("repair_studio_objective_issues", input, (result) => result.requiresNativeRerender ? `Fixed ${result.fixedIssueIds.length} bounded native issue${result.fixedIssueIds.length === 1 ? "" : "s"} on slide ${result.slideNumber}; ${result.deferredIssueIds.length} issue${result.deferredIssueIds.length === 1 ? "" : "s"} remain deferred. Rebuild and inspect the new PowerPoint raster before judgment.` : `No safe deterministic geometry change was available on slide ${result.slideNumber}; follow the ${result.deferredIssueIds.length} deferred design route${result.deferredIssueIds.length === 1 ? "" : "s"}.`));

server.registerTool("record_studio_visual_critique", {
  title: "Record a bounded Studio visual-quality pass",
  description: "Record the AI's visual judgment for the exact Studio scene revision and PowerPoint raster returned by get_studio_slide_critique. Supply concrete source strengths, candidate improvements and regressions, a source/candidate/equivalent preference, and only visible issues after comparing original and export-result pixels. Source wins when the candidate is weaker. A ready verdict is withheld when native blocker or major issues remain, when the AI reports blocker/major issues, or when source is preferred. Passes are capped at three; unresolved pass 3 becomes hold for human review. Recording critique metadata does not change slide design geometry, save a project, or export PowerPoint.",
  inputSchema: {
    deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }), expectedSceneRevision: z.string().min(1).max(500), slideNumber: z.number().int().min(1), rasterSha256: z.string().length(64),
    verdict: z.enum(["ready", "revise", "hold"]), rationale: z.string().min(1).max(1_000),
    sourceComparison: z.object({ preferred: z.enum(["source", "candidate", "equivalent"]), sourceStrengths: z.array(z.string().min(1).max(500)).max(12), candidateImprovements: z.array(z.string().min(1).max(500)).max(12), candidateRegressions: z.array(z.string().min(1).max(500)).max(12) }),
    visualIssues: z.array(z.object({ category: z.enum(["alignment", "spacing", "hierarchy", "figure", "brand", "legibility", "consistency", "other"]), severity: z.enum(["blocker", "major", "minor"]), nodeIds: z.array(z.string().min(1).max(180)).max(30).default([]), message: z.string().min(1).max(1_000), recommendation: z.string().min(1).max(1_000), autoFixable: z.boolean().default(false) })).max(30).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("record_studio_visual_critique", input, (result) => `Recorded Studio visual pass ${result.review.pass}/3 as ${result.review.recordedVerdict}. ${result.review.recordedVerdict === "ready" ? "The exact slide revision is ready for human review." : result.review.recordedVerdict === "revise" ? "Use the issue ledger for one materially different refinement." : "The bounded loop is held for human review."}`));

server.registerTool("build_studio_presentation", {
  title: "Build the central Studio presentation",
  description: "Start or reuse one background build of every slide from the one persisted Studio Web Scene, including converted ORNL Template Pack artwork, into one editable PowerPoint candidate. The app stays usable and visibly reports compile, template, render, measurement, and hard-QA progress. Source-only slides, incomplete exact-content mapping, missing media, unsupported native internals, changed text/table content, PowerPoint text overflow, distorted protected marks, and non-authoritative rendering hold the build. Poll get_studio_presentation_build_status until ready; never call a started job finished. The project revision is crash-checkpointed automatically; this does not save or export a user-named file.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("build_studio_presentation", input, (result) => result.job.phase === "ready" ? `The exact central Studio result for ${result.deck.name} is already PowerPoint-validated and ready for review. No file was saved or exported.` : `Started background central build ${result.job.id} for ${result.deck.name}. Poll get_studio_presentation_build_status; do not call the deck ready yet.`));

server.registerTool("get_studio_presentation_build_status", {
  title: "Read central Studio build progress",
  description: "Read the live compile, ORNL templating, native rendering, native measurement, and production-QA phase for a background central Studio build without blocking the app. Supply either the returned job ID or the deck ID. A job is reviewable only when phase is ready; failed, canceled, or superseded jobs are not export candidates.",
  inputSchema: {
    jobId: z.string().min(1).max(120).optional(),
    deckId: z.string().min(1).max(120).optional(),
  },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_studio_presentation_build_status", input, (result) => `Central build ${result.job.id} is ${result.job.phase} at ${result.job.progressPercent}%. ${result.job.message}`));

server.registerTool("cancel_studio_presentation_build", {
  title: "Cancel a central Studio build",
  description: "Request cancellation of one background central Studio build. The saved JSON scene, immutable source, and any prior verified result remain intact. Presentation Studio discards the candidate at the next safe PowerPoint boundary.",
  inputSchema: { jobId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("cancel_studio_presentation_build", input, (result) => `Cancellation requested for central build ${result.job.id}. The saved scene and prior verified result remain intact.`));

server.registerTool("run_deck_qualification", {
  title: "Build the native deck qualification evidence bundle",
  description: "Reopen the exact current central Studio candidate and immutable source in Microsoft PowerPoint, export every slide from both as a private 2,200-pixel PNG, remeasure the candidate, and write a local evidence bundle with exact-content, table-structure, Aptos, overflow, off-slide, protected-template, and material-design-impact checks. The report routes slide-design failures to bounded MCP work, repeated engine defects to code regression work, and visual-asset needs to the governed concept queue. This does not apply a proposal, save the project, export the presentation to a user destination, or prove that the design is better; every candidate PNG still requires full-size visual review.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("run_deck_qualification", input, (result) => `Qualified all ${result.report.totals.slides} source/candidate slide pairs for ${result.deck.name}. Found ${result.report.issues.length} routed objective issue${result.report.issues.length === 1 ? "" : "s"}; full-size visual review remains required.`));

server.registerTool("get_deck_qualification", {
  title: "Read the latest deck qualification ledger",
  description: "Read the exact current central Studio scene's objective qualification checks, per-slide issue IDs, repair routing, PowerPoint provenance, and pending visual-review requirements without returning file bytes or images. Follow with get_qualification_slide for the source and candidate pixels of every slide; metadata alone cannot establish visual quality.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_qualification", input, (result) => `Read qualification ${result.report.id} for ${result.deck.name}: ${result.report.status}, ${result.report.issues.length} routed issue${result.report.issues.length === 1 ? "" : "s"}.`));

server.registerTool("get_qualification_slide", {
  title: "View one exact qualification slide image",
  description: "Return a full-resolution Microsoft PowerPoint-native PNG, a precise native-measurement issue crop, or a temporary diagnostic overlay from the latest source/candidate qualification pair. Request both clean full representations for every slide; use crops and overlays only to locate named issues, never as finished-slide artwork or as a substitute for full-size visual judgment.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), representation: z.enum(["source", "candidate"]).default("candidate"), view: z.enum(["full", "issue-crop", "diagnostic-overlay"]).default("full"), issueId: z.string().min(1).max(180).optional() },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_qualification_slide", input, (result) => `Viewed the exact ${result.representation} PowerPoint-native ${result.view.replaceAll("-", " ")} for slide ${result.slideNumber} of ${result.deck.name}.`));

server.registerTool("get_qualification_contact_sheet", {
  title: "View a qualification deck overview",
  description: "Return one page of up to 40 authoritative PowerPoint-native source or exact central-candidate thumbnails from the latest qualification. Use all pages to find cross-slide hierarchy, pacing, density, repetition, table, figure, and consistency outliers; then inspect every clean full slide pair before recording visual review.",
  inputSchema: { deckId: z.string().min(1).max(120), representation: z.enum(["source", "candidate"]).default("candidate"), page: z.number().int().min(1).default(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_qualification_contact_sheet", input, (result) => `Viewed ${result.representation} qualification overview page ${result.page} of ${result.pageCount}, slides ${result.firstSlideNumber}–${result.lastSlideNumber}.`));

server.registerTool("record_deck_qualification_review", {
  title: "Record revision-bound full-deck visual review",
  description: "Record visual verdicts for 1–40 exact source/candidate slide pairs after inspecting their clean full-resolution PowerPoint images. Every review must state source strengths, candidate improvements and regressions, and whether source, candidate, or equivalent is preferred. The source wins automatically when the candidate is weaker. Every verdict is bound to both raster hashes, the candidate SHA-256, and the scene revision. A ready verdict is withheld when objective issues, major visual findings, or a source-preferred comparison remain. Authorized-AI retries are capped at three before hold. This writes only private qualification metadata; it does not change slide design, save the project, export PowerPoint, or constitute formal ORNL approval.",
  inputSchema: {
    deckId: z.string().min(1).max(120), expectedSceneRevision: z.string().min(1).max(500), qualificationId: z.string().min(1).max(180), candidateSha256: z.string().length(64),
    reviews: z.array(z.object({
      slideNumber: z.number().int().min(1), sourceRasterSha256: z.string().length(64), candidateRasterSha256: z.string().length(64), verdict: z.enum(["ready", "revise", "hold"]), rationale: z.string().min(1).max(2_000),
      sourceComparison: z.object({
        preferred: z.enum(["source", "candidate", "equivalent"]),
        sourceStrengths: z.array(z.string().min(1).max(500)).max(12),
        candidateImprovements: z.array(z.string().min(1).max(500)).max(12),
        candidateRegressions: z.array(z.string().min(1).max(500)).max(12),
      }),
      findings: z.array(z.object({ category: z.enum(["hierarchy", "alignment", "spacing", "layout-balance", "table-quality", "figure-clarity", "template-fidelity", "deck-consistency", "source-intent", "other"]), severity: z.enum(["major", "minor"]), message: z.string().min(1).max(1_000), repairRoute: z.enum(["mcp-design", "engine-code", "image-concept", "human-review"]) })).max(12).default([]),
    })).min(1).max(40),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("record_deck_qualification_review", input, (result) => result.status === "review-complete" ? `Recorded exact-pixel visual review for all ${result.visualAcceptance.reviewedSlideCount} slides. The deck is ready for human draft review; this is not formal ORNL approval.` : `Recorded ${result.visualAcceptance.reviewedSlideCount} visual reviews; ${result.visualAcceptance.revisionSlideCount} need revision and ${result.visualAcceptance.heldSlideCount} are held.`));

server.registerTool("list_resources", {
  title: "List automatically shared project Resources",
  description: "List every embedded project Resource automatically shared when the person turns on the app's single AI access switch. Requires AI access. This inventory never returns original file bytes; compatible documents/data expose bounded extracted text, images expose a bounded preview, and unsupported formats expose metadata only.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("list_resources", {}, (result) => `Read metadata for all ${result.resources.length} project Resource${result.totalResourceCount === 1 ? "" : "s"} automatically shared by the active AI session.`));

server.registerTool("get_resource_text", {
  title: "Read automatically shared Resource text",
  description: "Read one bounded page of locally extracted text from a compatible document or data Resource automatically shared by the app's single AI access switch. Returns the embedded derivative hash, offsets, and truncation state so a model can ground a new presentation without reading the external original. Use exact excerpts from this result when creating slides; never invent missing source content.",
  inputSchema: { resourceId: z.string().min(1).max(180), offset: z.number().int().nonnegative().default(0), maximumCharacters: z.number().int().min(1_000).max(40_000).default(20_000) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_resource_text", input, (result) => `Read ${result.characterCount} of ${result.totalCharacterCount} extracted characters from ${result.resource.name}${result.nextOffset === undefined ? "" : `; continue at offset ${result.nextOffset}`}.`));

server.registerTool("get_resource_preview", {
  title: "View an automatically shared image Resource",
  description: "Return one bounded image preview for an embedded image Resource automatically shared by the app's single AI access switch. Use this for approved source imagery, visual references, or concept-only Image Gen drafts. The original file remains local and is never modified; metadata-only Resources do not expose pixels.",
  inputSchema: { resourceId: z.string().min(1).max(180) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_resource_preview", input, (result) => `Viewed the automatically shared ${result.resource.roleLabel} image Resource ${result.resource.name}. Treat concept-only pixels as visual direction, never as authority for text, logos, data, claims, or technical relationships.`));

server.registerTool("create_studio_presentation", {
  title: "Create a source-grounded native Studio presentation",
  description: "Create a brand-new editable 16:9 presentation in the one canonical native Studio JSON/HTML/CSS scene from source excerpts and image Resources automatically shared by the app's single AI access switch. The first slide must use an approved ORNL title layout; it is sacred, and Studio always applies it. For later slides, specify the communication archetype and omit recipe to let Studio select an exact compatible native ORNL layout contract or the controlled shared responsive recipe; provide recipe only for a deliberate override. Every slide and node retains Resource hashes and exact source excerpts. This creates an embedded editable PowerPoint source plus the central Studio scene, crash-checkpoints that project revision, and leaves it visible for review; it does not save a user-named project file or export PowerPoint to a user destination. Read all required Resource text and the Template Pack catalog first.",
  inputSchema: {
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    name: z.string().min(1).max(240),
    communicationJob: z.string().min(1).max(1_000),
    expression: z.enum(["restrained", "balanced", "expressive"]).default("balanced"),
    slides: z.array(z.object({
      title: z.string().min(1).max(300),
      subtitle: z.string().min(1).max(600).optional(),
      body: z.array(z.string().min(1).max(1_500)).max(12).default([]),
      archetype: z.enum(["cover", "section", "assertion-evidence", "text-led", "hero-figure", "comparison", "image-series", "portrait-series", "table", "data-visualization", "process-flow", "technical-diagram", "conclusion"]).optional(),
      recipe: z.enum(["ornl-title-content", "ornl-title-two-column", "ornl-title-card-grid", "ornl-title-metric-grid", "ornl-title-table", "ornl-title-figure-grid", "ornl-title-objective-columns", "ornl-title-steps-evidence", "ornl-title-labeled-figure-grid", "ornl-title-image-series", "ornl-title-question-diagram", "ornl-title-challenges-evidence", "ornl-title-process-flow", "template-layout"]).optional(),
      layoutId: z.string().min(1).max(180).optional(),
      imageResourceIds: z.array(z.string().min(1).max(180)).max(14).default([]),
      table: z.object({ headers: z.array(z.string().min(1).max(500)).min(1).max(12), rows: z.array(z.array(z.string().max(2_000)).min(1).max(12)).max(40) }).optional(),
      sourceReferences: z.array(z.object({ resourceId: z.string().min(1).max(180), exactExcerpt: z.string().min(1).max(20_000) })).min(1).max(12),
      rationale: z.string().min(1).max(1_000).optional(),
    })).min(1).max(80),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("create_studio_presentation", input, (result) => `Created ${result.slideCount} source-grounded ORNL slides in the native Studio scene for ${result.deck.name}. The central design is open for review; nothing was saved or exported.`));

server.registerTool("create_studio_visual_need", {
  title: "Create a governed visual-direction brief",
  description: "Add one model-independent visual need to a non-protected Studio slide when layout art direction, a figure concept, image treatment, supporting visual, or diagram-rebuild concept could materially improve communication. The default brief exposes abstract structure only and contains no source wording or pixels. Exact-content disclosure requires an explicitly bounded approved summary. This creates a local prompt package but does not call an image model, attach pixels, redesign, build, save, or export anything.",
  inputSchema: {
    deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }), slideNumber: z.number().int().min(1),
    type: z.enum(["layout-concept", "figure-concept", "image-treatment", "supporting-visual", "diagram-rebuild"]),
    reason: z.string().min(1).max(1_000), communicationJob: z.string().min(1).max(1_000),
    expression: z.enum(["restrained", "balanced", "expressive"]).default("balanced"),
    approvedInfluences: z.array(z.enum(["composition", "visual-hierarchy", "negative-space", "color-balance", "figure-concept", "image-treatment", "visual-rhythm"])).min(1).max(7).optional(),
    disclosurePolicy: z.enum(["abstract-structure-only", "exact-content-approved"]).default("abstract-structure-only"),
    approvedContentSummary: z.string().min(1).max(800).optional(),
    targetSlot: z.object({ role: z.enum(["whole-slide", "primary-visual", "supporting-evidence", "figure", "background-treatment"]).optional(), aspectRatio: z.enum(["16:9", "4:3", "1:1", "free"]).optional(), placementNotes: z.string().min(1).max(500).optional() }).optional(),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("create_studio_visual_need", input, (result) => `Created a ${result.need.type.replaceAll("-", " ")} brief for Studio slide ${result.slideNumber}. Disclosure is ${result.need.disclosurePolicy}; no image model was called and no source or presentation file was changed.`));

server.registerTool("list_studio_visual_needs", {
  title: "List the Studio visual-needs queue",
  description: "List local visual-direction requests, lifecycle state, source-content binding, target slot, and whether an approved concept is linked. This returns no slide text, source pixels, concept pixels, prompts, or presentation bytes.",
  inputSchema: { deckId: z.string().min(1).max(120).optional() },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("list_studio_visual_needs", input, (result) => `Read ${result.needs.length} visual need${result.needs.length === 1 ? "" : "s"} from the local Studio queue.`));

server.registerTool("get_studio_visual_need_brief", {
  title: "Read a governed visual-direction prompt package",
  description: "Read one source-hash-bound visual brief with abstract structure inventory, approved ORNL expression, target slot, positive prompt, negative prompt, and disclosure boundary. The default prompt never contains source wording or source pixels. Use any image generator only within that boundary, then import its result as an image Resource and attach it as concept-only art direction.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), visualNeedId: z.string().min(1).max(180) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_studio_visual_need_brief", input, (result) => `Read the governed ${result.need.type.replaceAll("-", " ")} brief for slide ${result.slideNumber}. Follow its ${result.need.disclosurePolicy} boundary before using any external image model.`));

server.registerTool("hold_studio_visual_need", {
  title: "Hold a Studio visual need",
  description: "Move one visual need out of the active queue while preserving its brief and provenance. Use this when a concept is unnecessary, unsafe, ambiguous, or should wait for human/content-owner direction. This does not delete Resources, concepts, slides, or source files.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }), slideNumber: z.number().int().min(1), visualNeedId: z.string().min(1).max(180), note: z.string().min(1).max(1_000) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("hold_studio_visual_need", input, (result) => `Held visual need ${result.visualNeedId} on slide ${result.slideNumber}. Its local brief remains available and no source file changed.`));

server.registerTool("attach_studio_concept_reference", {
  title: "Attach a concept-only visual reference to a Studio slide",
  description: "Attach an embedded image Resource automatically preview-shared by the active AI access switch as non-authoritative art direction for one non-protected Studio slide. Record only the approved characteristics to follow and a normalized composition blueprint. Generated text, logos, data, claims, and technical details are always untrusted; exact source content and approved assets remain authoritative. This changes design input only—it does not trace pixels, redesign, build, save, or export the slide.",
  inputSchema: {
    deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }), slideNumber: z.number().int().min(1), resourceId: z.string().min(1).max(180),
    origin: z.enum(["imagegen", "human-reference", "other"]),
    approvedInfluences: z.array(z.enum(["composition", "visual-hierarchy", "negative-space", "color-balance", "figure-concept", "image-treatment", "visual-rhythm"])).min(1).max(7),
    blueprint: z.object({
      summary: z.string().min(1).max(1_000),
      zones: z.array(z.object({ id: z.string().min(1).max(120), role: z.enum(["title", "primary-visual", "supporting-evidence", "caption", "footer-safe", "other"]), x: z.number().min(0).max(1), y: z.number().min(0).max(1), width: z.number().positive().max(1), height: z.number().positive().max(1) })).max(20).default([]),
      styleNotes: z.array(z.string().min(1).max(500)).max(20).default([]),
      reconstructionNotes: z.array(z.string().min(1).max(500)).max(20).default([]),
    }),
    provenance: z.object({ model: z.string().min(1).max(180).optional(), promptSummary: z.string().min(1).max(1_000).optional(), generatedAt: z.string().datetime({ offset: true }).optional() }).optional(),
    visualNeedId: z.string().min(1).max(180).optional(),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("attach_studio_concept_reference", input, (result) => `Attached ${result.reference.origin} concept ${result.reference.id} to Studio slide ${result.slideNumber}. Use it for ${result.reference.approvedInfluences.join(", ")} only; reconstruct with exact source content, then build and inspect the native PowerPoint result.`));

server.registerTool("get_studio_concept_reference", {
  title: "View a slide's concept reference and reconstruction blueprint",
  description: "Return the bounded concept-only image plus its approved visual influences, normalized blueprint, source-content binding, provenance, and mandatory untrusted-element list. Use it alongside get_studio_web_scene and the original PowerPoint render; never copy generated wording, logos, data, claims, or technical details into the editable reconstruction.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), referenceId: z.string().min(1).max(180) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_studio_concept_reference", input, (result) => `Viewed concept-only reference ${result.reference.id} for slide ${result.slideNumber}. The source PowerPoint remains authoritative for all content and technical meaning.`));

server.registerTool("reconstruct_studio_concept", {
  title: "Reconstruct approved concept zones as editable Studio content",
  description: "Convert one linked concept-only reference's normalized semantic zones into a material editable Studio composition using exact source-bound text, tables, media, and technical evidence. This is the deterministic bridge from concept art direction to the central design scene: it never traces the raster, copies generated typography, uses generated claims, changes the sacred ORNL title slide, builds PowerPoint, saves, or exports. A visual need advances only when the result changes actual layout or figure treatment rather than merely shrinking type.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    referenceId: z.string().min(1).max(180),
    recipe: z.enum(["ornl-title-content", "ornl-title-two-column", "ornl-title-card-grid", "ornl-title-metric-grid", "ornl-title-table", "ornl-title-figure-grid", "ornl-title-objective-columns", "ornl-title-steps-evidence", "ornl-title-labeled-figure-grid", "ornl-title-image-series", "ornl-title-question-diagram", "ornl-title-challenges-evidence", "ornl-title-process-flow"]).optional(),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("reconstruct_studio_concept", input, (result) => `Reconstructed concept ${result.referenceId} into ${result.mappedNodeIds.length} editable source-bound element${result.mappedNodeIds.length === 1 ? "" : "s"} on Studio slide ${result.slideNumber}. Build and inspect the PowerPoint-native result before review; nothing was saved or exported.`));

server.registerTool("remove_studio_concept_reference", {
  title: "Detach a concept reference from a Studio slide",
  description: "Remove one concept-only reference from the slide's design inputs without deleting its embedded Resource or changing the source PowerPoint. The concept remains available in Resources until the person removes it separately.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }), slideNumber: z.number().int().min(1), referenceId: z.string().min(1).max(180) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("remove_studio_concept_reference", input, (result) => `Detached concept ${result.referenceId} from Studio slide ${result.slideNumber}; the embedded Resource and source PowerPoint remain unchanged.`));

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

server.registerTool("get_studio_composition_plan", {
  title: "Plan one slide as an ORNL communication archetype",
  description: "Classify one slide by its communication job—cover, assertion-evidence, text-led, hero figure, comparison, image series, portrait series, table, data visualization, process flow, technical diagram, conclusion, or source preserve—then return the safest high-level Studio recipe, compatible native ORNL layout contract, explicit preserve/polish/recompose/rebuild-figure intervention, the exact Presentation Design Standard version, and deterministic table-capacity guidance governing the plan. Use this before staging design. When tableCapacity.required is true, stage the chosen table design and then call plan_studio_table_continuation with the returned recommendedMaximumBodyRowsPerSlide before preview or build; never present the known-over-capacity one-slide result as a candidate. The planner prefers an exact approved native layout only when all source relationships fit; otherwise it chooses a shared responsive Studio archetype on the neutral native ORNL base. Source always wins when a candidate is not visibly stronger. This is a design plan, not visual approval.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    slideNumber: z.number().int().min(1),
    archetype: z.enum(["cover", "section", "assertion-evidence", "text-led", "hero-figure", "comparison", "image-series", "portrait-series", "table", "data-visualization", "process-flow", "technical-diagram", "conclusion", "source-preserve"]).optional(),
  },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_studio_composition_plan", input, (result) => `Planned slide ${result.slide.number} of ${result.deck.name} as ${result.plan.archetype} using ${result.plan.strategy}.`));

server.registerTool("get_slide_design_work_order", {
  title: "Read one versioned AI slide-design work order",
  description: "Assemble the authoritative Current PowerPoint evidence, exact locked copy, hybrid scene objects and allowed operations, slide findings, submitted comments, ORNL rules, and ranked Template Pack layouts into one revision-bound work order. Call this before designing a slide; it is the primary Inspect and Diagnose input for the iterative visual-design loop.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_slide_design_work_order", input, (result) => `Built image-bearing design work order ${result.revision} for slide ${result.slide.number} of ${result.deck.name}. The attached PNG is an authoritative 2,200-pixel Microsoft PowerPoint render.`));

server.registerTool("get_slide_inspection_packet", {
  title: "Inspect a slide with native pixels, crops, measurements, and metrics",
  description: "Return one revision-bound source-bound cleanup inspection packet containing the exact design work order, a 2,200-pixel PowerPoint-native full-slide PNG, readable title/table/text crops, a deterministic crop overlay, native rendered-text and cell measurements, objective design metrics, and a Found issues repair ledger. Current inspection records issues and the exact original-intent reference; Proposal inspection enters Rechecking original intent for a staged cleanup proposal. For a fresh central Studio composition, use preview_studio_fresh_composition followed by get_studio_slide_critique instead of requesting Proposal here. Pixels guide gestalt, PowerPoint supplies measurements, and deterministic solvers supply exact coordinates.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), representation: z.enum(["current", "proposal"]).default("current") },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImages("get_slide_inspection_packet", input, (result) => `Built ${result.representation} inspection packet ${result.revision} with ${result.images.length} PowerPoint-native visual evidence images for slide ${result.slide.number} of ${result.deck.name}.`));

server.registerTool("get_deck_design_work_order", {
  title: "Read the representative deck-design qualification set",
  description: "Select up to eight high-complexity representative slides, one for each communication archetype present, and return a complete versioned work order for each plus archetype coverage and the deck-wide semantic table-color map. Qualify this bounded set before propagating any layout pattern across the full deck. Exact source content is included, so visible AI session access is required.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_design_work_order", input, (result) => `Built ${result.workOrders.length} communication-archetype qualification work orders for ${result.deck.name}.`));

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
  description: "After reading the Proposal inspection packet, record whether the native PowerPoint draft is visually better, needs another semantic revision, or should be rejected. Recheck the proposal against the original human slide's message: exact wording, source visual identity or disclosed verified replacement, meaning-bearing labels/values, and arrows/sequence/causality/grouping. The exact inspection revision, raster hashes, objective metric changes, intent review, rationale, and attempt number are persisted. A requested better verdict is withheld when metrics regress, pixels are unchanged, or intent remains unverified. Automatic AI revision is capped at three attempts; attempt three rejects an unresolved draft. This never applies, saves, exports, or overwrites content.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    proposalId: z.string().min(1).max(120),
    slideNumber: z.number().int().min(1),
    inspectionRevision: z.string().min(1).max(500),
    verdict: z.enum(["better", "revise", "reject"]),
    rationale: z.string().min(1).max(1_000),
    intentReview: z.object({
      status: z.enum(["pass", "needs-review"]),
      exactTextPreserved: z.boolean(),
      sourceVisualsPreserved: z.boolean(),
      relationshipsPreserved: z.enum(["yes", "not-applicable", "unverified"]),
      summary: z.string().min(1).max(1_000),
    }),
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
  description: "For an ornl-studio target only, update the one canonical Studio presentation by making a substantive layout decision from a communication archetype, then compile it through fresh-composition. Call get_studio_composition_plan first and echo its exact designStandardVersion; stale guidance is rejected before any scene mutation. You may provide archetype here and omit recipe to let the planner choose a compatible native ORNL layout or shared responsive Studio archetype. Converted layouts remain vocabulary in the same scene, not alternate proposals. This tool rejects source-template-cleanup decks because sponsor/custom masters and layouts must instead use source-bound proposals unless the user explicitly selected ORNL cross-template conversion. HARD RULE: an existing populated ORNL title slide is sacred and may only use source. Preserve meaning-bearing UI, code, labels, values, arrows, sequence, causality, and repeated source relationships. For image-heading-evidence groups, never flatten the groups into stacked text and tiny thumbnails; likewise, do not flatten technical relationships into unrelated text. Optional nodeFrames are final refinements, not a layout substitute. Judge the compiled result through PowerPoint-native visual QA; an intermediate recipe canvas is not reviewable output. A technical build is only a candidate and cannot become exportable until its exact raster passes revision-bound visual review. The project revision is crash-checkpointed automatically; no user-named project file or export is created or overwritten.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    designStandardVersion: z.string().min(1).max(160),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    slideNumber: z.number().int().min(1),
    archetype: z.enum(["cover", "section", "assertion-evidence", "text-led", "hero-figure", "comparison", "image-series", "portrait-series", "table", "data-visualization", "process-flow", "technical-diagram", "conclusion", "source-preserve"]).optional(),
    interventionLevel: z.enum(["preserve", "polish", "recompose", "rebuild-figure"]).optional(),
    interventionRationale: z.string().min(1).max(1_000).optional(),
    recipe: z.enum(["source", "ornl-title-content", "ornl-title-two-column", "ornl-title-card-grid", "ornl-title-metric-grid", "ornl-title-table", "ornl-title-figure-grid", "ornl-title-objective-columns", "ornl-title-steps-evidence", "ornl-title-labeled-figure-grid", "ornl-title-image-series", "ornl-title-question-diagram", "ornl-title-challenges-evidence", "ornl-title-process-flow", "template-layout"]).optional(),
    compilerMode: z.literal("fresh-composition").default("fresh-composition"),
    layoutId: z.string().min(1).max(120).optional(),
    rationale: z.string().min(1).max(1000).optional(),
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
    figureTreatments: z.array(z.object({
      id: z.string().min(1).max(180).optional(),
      nodeIds: z.array(z.string().min(1).max(180)).min(1).max(200),
      mode: z.enum(["preserve-as-unit", "preserve-and-frame", "hybrid-rebuild", "redraw-candidate"]),
      verificationStatus: z.enum(["source-locked", "needs-content-review", "verified"]),
      intentSummary: z.string().min(1).max(1_000),
      informationInventory: z.array(z.string().min(1).max(500)).min(1).max(40),
      invariants: z.array(z.string().min(1).max(500)).min(1).max(40),
      rationale: z.string().min(1).max(1_000),
      relationships: z.array(z.object({
        fromNodeId: z.string().min(1).max(180),
        toNodeId: z.string().min(1).max(180),
        kind: z.enum(["caption-for", "label-for", "callout-for", "connects-from", "connects-to", "contained-by"]),
      })).max(80).default([]),
      groupFrame: z.object({
        xInches: z.number().min(0).max(20),
        yInches: z.number().min(0).max(20),
        widthInches: z.number().min(.1).max(20),
        heightInches: z.number().min(.1).max(20),
        rotation: z.number().min(-360).max(360).default(0),
      }).optional(),
      focalPoint: z.object({ x: z.number().min(0).max(1), y: z.number().min(0).max(1) }).optional(),
      crop: z.object({ left: z.number().min(0).max(.99), top: z.number().min(0).max(.99), right: z.number().min(0).max(.99), bottom: z.number().min(0).max(.99) }).optional(),
      relationshipPolicy: z.enum(["preserve-internal", "reflow-annotations", "editable-diagram"]).optional(),
      lockAspectRatio: z.boolean().optional(),
    })).max(40).default([]),
    visualNeedIds: z.array(z.string().min(1).max(180)).max(8).default([]),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_studio_web_design", input, (result) => `Recomposed slide ${result.slide.number} with ${result.slide.recipe} in the one canonical Studio presentation. Build its current revision and judge the PowerPoint-native result; nothing was saved, exported, or overwritten.`));

server.registerTool("refine_studio_layout", {
  title: "Refine Studio layout with optical constraints",
  description: "Apply 1–20 high-level layout constraints to the canonical Studio HTML/CSS scene without guessing PowerPoint coordinates. Align uses structural edges or PowerPoint-native rendered text bounds for optical-left/optical-top when the exact current slide build exists; distribute creates equal gaps; snap-to-grid uses the deck rhythm; fit-safe-region moves an intact group minimally. Supply relationship-preserving groups for figures, diagrams, captions, cards, and labels so their internal geometry moves together. The solver rejects new overlaps, off-canvas results, locked objects, and infeasible groups. Build the exact slide afterward and repeat an optical pass when evidenceAuthority is scene-estimate. The sacred ORNL title slide cannot be refined.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    constraints: z.array(z.object({
      kind: z.enum(["align", "distribute", "snap-to-grid", "fit-safe-region"]),
      mode: z.enum(["left", "optical-left", "center", "right", "top", "optical-top", "middle", "bottom", "horizontal-equal-gap", "vertical-equal-gap", "both"]),
      nodeIds: z.array(z.string().min(1).max(180)).min(1).max(60),
      groups: z.array(z.array(z.string().min(1).max(180)).min(1).max(30)).min(1).max(30).optional(),
      anchorNodeId: z.string().min(1).max(180).optional(),
      gridPt: z.number().min(1).max(72).optional(),
      rationale: z.string().min(1).max(1_000),
    })).min(1).max(20),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("refine_studio_layout", input, (result) => `Applied ${result.constraints.length} high-level Studio layout constraint${result.constraints.length === 1 ? "" : "s"} to slide ${result.slideNumber} using ${result.evidenceAuthority} evidence. Build and inspect the exact current revision; nothing was saved or exported.`));

server.registerTool("refine_studio_table", {
  title: "Refine one source-bound Studio table",
  description: "Edit one native Studio table component by stable node and cell IDs. Set complete column widths and row heights in inches, header-row count, global or per-cell/per-edge border treatment, default padding, and bounded cell styles while preserving every source cell's exact text, order, merged spans, source-significant semantic color role, and native editability. Use edge rules for deliberate header separators, totals, and semantic group boundaries instead of forcing a full grid. Use 8.5–9.5 pt table text only as a bounded source-equivalent exception for an extreme exact-content continuation after allocating the complete ORNL table region; ordinary dense technical tables remain at least 10 pt. Use the table IDs and current dimensions from get_studio_web_scene; do not infer missing cells, overwrite a semantic fill, or alter technical meaning. Build the exact slide afterward and inspect PowerPoint-native cell measurements. The project revision is crash-checkpointed automatically; no user-named project file or export is created.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    tableNodeId: z.string().min(1).max(180),
    columnWidthsInches: z.array(z.number().min(.35).max(20)).max(30).optional(),
    rowHeightsInches: z.array(z.number().min(.18).max(10)).max(100).optional(),
    headerRows: z.number().int().min(0).max(20).optional(),
    borderMode: z.enum(["none", "subtle", "full"]).optional(),
    borderColor: z.string().regex(/^#[0-9a-f]{6}$/i).optional(),
    borderWidthPt: z.number().min(0).max(6).optional(),
    defaultPaddingPt: z.number().min(0).max(36).optional(),
    cellStyles: z.array(z.object({
      cellId: z.string().min(1).max(180),
      fill: z.string().regex(/^#[0-9a-f]{6}$/i).optional(),
      color: z.string().regex(/^#[0-9a-f]{6}$/i).optional(),
      fontSizePt: z.number().min(8.5).max(40).optional(),
      fontWeight: z.union([z.literal(400), z.literal(600), z.literal(700)]).optional(),
      textAlign: z.enum(["left", "center", "right"]).optional(),
      verticalAlign: z.enum(["top", "middle", "bottom"]).optional(),
      paddingPt: z.number().min(0).max(36).optional(),
      borders: z.object({
        top: z.object({ type: z.enum(["none", "solid", "dash"]), color: z.string().regex(/^#[0-9a-f]{6}$/i), widthPt: z.number().min(0).max(6) }).optional(),
        right: z.object({ type: z.enum(["none", "solid", "dash"]), color: z.string().regex(/^#[0-9a-f]{6}$/i), widthPt: z.number().min(0).max(6) }).optional(),
        bottom: z.object({ type: z.enum(["none", "solid", "dash"]), color: z.string().regex(/^#[0-9a-f]{6}$/i), widthPt: z.number().min(0).max(6) }).optional(),
        left: z.object({ type: z.enum(["none", "solid", "dash"]), color: z.string().regex(/^#[0-9a-f]{6}$/i), widthPt: z.number().min(0).max(6) }).optional(),
      }).optional(),
    })).max(200).default([]),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("refine_studio_table", input, (result) => `Refined ${result.table.rows} × ${result.table.columns} source-bound table ${result.table.nodeId} on slide ${result.slideNumber}. Exact cell copy, merge topology, and semantic roles remain locked; build and inspect the PowerPoint-native result before review.`));

server.registerTool("publish_studio_table_exemplar", {
  title: "Publish one approved Studio table exemplar",
  description: "Adopt the current source-bound table's approved visual treatment as a reusable deck definition and apply it to structurally compatible tables. Compatibility requires the same column count, header-row count, header merge pattern, and body merge pattern. The operation propagates column proportions, borders, padding, typography, alignment, and nonsemantic fills only; it never copies cell content, changes merged topology, or overwrites source-significant semantic colors. Build every affected slide and inspect native PowerPoint table measurements afterward. The project revision is crash-checkpointed automatically; no user-named project file or export is created.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    tableNodeId: z.string().min(1).max(180),
    name: z.string().min(1).max(120).optional(),
    targetSlideNumbers: z.array(z.number().int().min(1)).max(200).optional(),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("publish_studio_table_exemplar", input, (result) => `Published ${result.definition.name} and updated ${result.affectedTableNodeIds.length} structurally compatible native table${result.affectedTableNodeIds.length === 1 ? "" : "s"}. Exact content, merge topology, and semantic fills remain locked; rebuild the affected slides.`));

server.registerTool("plan_studio_table_continuation", {
  title: "Plan a merge-safe Studio table continuation",
  description: "Create or replace a revision-bound continuation plan for one source-bound native table. Studio repeats identified header rows and partitions body rows only at boundaries that do not split a merged cell. It records blockers instead of hiding rows, dropping cells, rasterizing the table, or shrinking below the ORNL minimum. A ready plan is automatically materialized as editable output slides by the next slide or Build-all PowerPoint compilation and validated against the explicit source-to-output map. This operation updates only the local Studio design; nothing is saved to disk or exported.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    tableNodeId: z.string().min(1).max(180),
    maximumBodyRowsPerSlide: z.number().int().min(1).max(40).default(8),
    rationale: z.string().min(1).max(1_000).optional(),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("plan_studio_table_continuation", input, (result) => result.plan.status === "ready" ? `Planned ${result.plan.segments.length} merge-safe continuation slides with ${result.plan.headerRows} repeated header row${result.plan.headerRows === 1 ? "" : "s"}. The next PowerPoint build materializes and validates them; nothing has been saved or exported.` : `Continuation planning is held: ${result.plan.blockers.join(" ")}`));

server.registerTool("clear_studio_table_continuation", {
  title: "Clear one Studio table continuation plan",
  description: "Remove a stored continuation plan from one table without changing the source-bound table, its content, or its visual design. The project revision is crash-checkpointed automatically; no user-named project file or export is created.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    tableNodeId: z.string().min(1).max(180),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("clear_studio_table_continuation", input, () => "Cleared the stored continuation plan. The native source-bound table remains unchanged."));

server.registerTool("author_studio_connector", {
  title: "Author one verified editable diagram connector",
  description: "Bind one existing Studio connector to two stable endpoint nodes inside the same verified editable-diagram figure treatment. Choose explicit attachment sides, line treatment, and arrowheads. Studio rejects guessed topology, endpoints outside the verified figure, source-locked figures, and stale scene revisions. The connector remains editable in PowerPoint and follows its endpoint nodes on the Studio canvas. Build and inspect the PowerPoint-native result afterward; nothing is saved or exported.",
  inputSchema: {
    deckId: z.string().min(1).max(120),
    expectedUpdatedAt: z.string().datetime({ offset: true }),
    expectedSceneRevision: z.string().min(1).max(500),
    slideNumber: z.number().int().min(1),
    connectorNodeId: z.string().min(1).max(180),
    fromNodeId: z.string().min(1).max(180),
    toNodeId: z.string().min(1).max(180),
    fromSide: z.enum(["top", "right", "bottom", "left", "center"]),
    toSide: z.enum(["top", "right", "bottom", "left", "center"]),
    stroke: z.string().regex(/^#[0-9a-f]{6}$/i).default("#00662C"),
    widthPt: z.number().min(.25).max(8).default(1.5),
    dash: z.enum(["solid", "dash", "dashDot"]).default("solid"),
    beginArrow: z.enum(["none", "arrow", "diamond", "oval", "stealth", "triangle"]).default("none"),
    endArrow: z.enum(["none", "arrow", "diamond", "oval", "stealth", "triangle"]).default("triangle"),
    addressedThreadIds: z.array(z.string().min(1).max(120)).max(40).default([]),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("author_studio_connector", input, (result) => `Bound verified editable connector ${result.connector.nodeId} from ${result.connector.fromNodeId} to ${result.connector.toNodeId} on slide ${result.slideNumber}. Build and inspect the PowerPoint-native result before review.`));

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
