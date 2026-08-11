import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { PresentationAppClient, PresentationAppUnavailableError } from "./presentation-app-client.mjs";

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

server.registerTool("stage_font_cleanup", {
  title: "Stage conservative font cleanup",
  description: "Stage supported legacy-font mappings for visible human review in the app. The correct ORNL target must already be confirmed by a person. Pass the exact updatedAt returned by list_decks. This does not apply, save, or export changes.",
  inputSchema: { deckId: z.string().min(1).max(120), expectedUpdatedAt: z.string().min(1).max(80) },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
}, (input) => call("stage_font_cleanup", input, (result) => `Staged ${result.proposal.summary} for human review. Nothing was applied, saved, or exported.`));

await server.connect(new StdioServerTransport());
