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

server.registerTool("list_resources", {
  title: "List authorized project Resources",
  description: "List metadata for project Resources that a person explicitly shared for the current app session. Requires AI session access. Never returns original file bytes or extracted document text.",
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, () => call("list_resources", {}, (result) => `Read metadata for ${result.resources.length} of ${result.totalResourceCount} project Resource${result.totalResourceCount === 1 ? "" : "s"} authorized for this session.`));

server.registerTool("get_deck_audit", {
  title: "Read a deck audit",
  description: "Read deterministic template, font, object, comments, support, and production findings for one audited deck. Requires AI session access. Visible slide text and Resource bytes are intentionally omitted.",
  inputSchema: { deckId: z.string().min(1).max(120) },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => call("get_deck_audit", input, (result) => `Read the audit for ${result.deck.name} with ${result.audit.findings.length} findings.`));

server.registerTool("get_slide_design_context", {
  title: "Read bounded slide design context",
  description: "Read the exact approved text plus typography and object inventory for up to 10 consecutive slides so the model can make layout decisions without rewriting content. Requires visible AI session access in Presentation Studio. This is semantic design context, not a rendered canvas; use slide-render tools when available for visual judgment.",
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

server.registerTool("get_slide_render", {
  title: "View a current or proposed slide design",
  description: "Return a bounded revision-labeled image of one current or proposal slide using Presentation Studio's local OOXML preview renderer. Compare both representations before recommending acceptance. The exported native PPTX and an independent PowerPoint render remain the final export-fidelity authority.",
  inputSchema: { deckId: z.string().min(1).max(120), slideNumber: z.number().int().min(1), representation: z.enum(["current", "proposal"]).default("current") },
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, (input) => callImage("get_slide_render", input, (result) => `Rendered ${result.representation} slide ${result.slide.number} of ${result.deck.name} at ${result.width} × ${result.height}.`));

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
  description: "Review every slide, stage supported exact-content typography, high-confidence text-box alignment, and native-table improvements, and record explicit approved-as-is or needs-review dispositions for the rest. Complex tables and semantic colors become exceptions instead of being flattened. This creates a reversible proposal with Current/Proposal renders; it does not apply, save, export, or overwrite anything.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().datetime({ offset: true }) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_designer_cleanup", input, (result) => `Staged a deck-wide designer proposal covering ${result.proposal.slideDispositions.length} slides. Nothing was applied, saved, or exported.`));

await server.connect(new StdioServerTransport());
