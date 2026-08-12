import assert from "node:assert/strict";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

test("standard STDIO MCP server advertises bounded audit, Resource, and proposal tools", async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const transport = new StdioClientTransport({ command: process.execPath, args: [path.join(root, "mcp", "server.mjs")] });
  const client = new Client({ name: "presentation-studio-test", version: "0.1.0" });
  try {
    await client.connect(transport);
    const result = await client.listTools();
    assert.deepEqual(result.tools.map((tool) => tool.name).sort(), ["get_app_status", "get_cleanup_rule_profile", "get_deck_audit", "get_design_contract", "get_design_thread", "get_slide_design_context", "get_slide_render", "list_decks", "list_design_threads", "list_resources", "stage_designer_cleanup", "stage_font_cleanup"]);
    const contractTool = result.tools.find((tool) => tool.name === "get_design_contract");
    assert.match(contractTool?.description ?? "", /improving every slide/i);
    const contract = await client.callTool({ name: "get_design_contract", arguments: {} });
    assert.equal(contract.isError, undefined);
    assert.equal((contract.structuredContent as { defaultMode?: { name?: string } })?.defaultMode?.name, "designer-cleanup");
    assert.match(JSON.stringify(contract.structuredContent), /independently rendered export/i);
    assert.match(JSON.stringify(contract.structuredContent), /design thread/i);
    assert.match(JSON.stringify(contract.structuredContent), /stable object, text range, table cell\/range, or normalized region/i);
    assert.match(JSON.stringify(contract.structuredContent), /measure every cell/i);
    assert.match(JSON.stringify(contract.structuredContent), /continuation slide/i);
    assert.match(JSON.stringify(contract.structuredContent), /reversible transaction history/i);
    assert.equal((contract.structuredContent as { defaults?: { slide?: { aspectRatio?: string }; typography?: { family?: string } } })?.defaults?.slide?.aspectRatio, "16:9");
    assert.equal((contract.structuredContent as { defaults?: { typography?: { family?: string } } })?.defaults?.typography?.family, "Aptos");
    assert.equal((contract.structuredContent as { tableProfile?: { cellPaddingPt?: { left?: number } } })?.tableProfile?.cellPaddingPt?.left, 6);
    const stageTool = result.tools.find((tool) => tool.name === "stage_font_cleanup");
    assert.equal(stageTool?.annotations?.destructiveHint, false);
    assert.equal(stageTool?.annotations?.readOnlyHint, false);
    const designerTool = result.tools.find((tool) => tool.name === "stage_designer_cleanup");
    assert.match(designerTool?.description ?? "", /every slide/i);
    assert.match(designerTool?.description ?? "", /semantic colors/i);
    assert.match(designerTool?.description ?? "", /text-box alignment/i);
    assert.equal(designerTool?.annotations?.destructiveHint, false);
  } finally {
    await client.close();
  }
});
