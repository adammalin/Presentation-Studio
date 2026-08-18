import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";

test("Codex MCP configuration is automatic, bounded, and idempotent", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "presentation studio codex config "));
  const target = path.join(root, "config.toml");
  try {
    writeFileSync(target, "model = \"gpt-test\"\n\n[mcp_servers.existing]\ncommand = \"existing\"\n");
    const script = path.resolve("scripts/configure-mcp.mjs");
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const result = spawnSync(process.execPath, [script, "--codex", target], { encoding: "utf8" });
      assert.equal(result.status, 0, result.stderr);
    }
    const configured = readFileSync(target, "utf8");
    assert.match(configured, /\[mcp_servers\.existing\]/);
    assert.match(configured, /\[mcp_servers\.presentation_studio\]/);
    assert.match(configured, /mcp\/server\.mjs/);
    assert.equal(configured.match(/BEGIN PRESENTATION STUDIO MCP/g)?.length, 1);
    assert.equal(configured.match(/END PRESENTATION STUDIO MCP/g)?.length, 1);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
