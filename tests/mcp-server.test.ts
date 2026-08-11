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
    assert.deepEqual(result.tools.map((tool) => tool.name).sort(), ["get_app_status", "get_deck_audit", "list_decks", "list_resources", "stage_font_cleanup"]);
    const stageTool = result.tools.find((tool) => tool.name === "stage_font_cleanup");
    assert.equal(stageTool?.annotations?.destructiveHint, false);
    assert.equal(stageTool?.annotations?.readOnlyHint, false);
  } finally {
    await client.close();
  }
});
