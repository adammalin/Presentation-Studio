import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (relativePath: string) => fs.readFileSync(path.join(root, relativePath), "utf8");

test("portable quality workflow cannot claim or upload native client evidence", () => {
  const workflow = read(".github/workflows/quality.yml");
  assert.match(workflow, /npm run quality -- --ci/);
  assert.doesNotMatch(workflow, /self-hosted|quality:native|upload-artifact|qualify:deck/);
  assert.match(workflow, /actions\/checkout@[0-9a-f]{40}/);
  assert.match(workflow, /actions\/setup-node@[0-9a-f]{40}/);
});

test("native workflow is manual, self-hosted, synthetic, and keeps evidence local", () => {
  const workflow = read(".github/workflows/native-powerpoint-qualification.yml");
  assert.match(workflow, /workflow_dispatch/);
  assert.match(workflow, /self-hosted/);
  assert.match(workflow, /presentation-studio-powerpoint/);
  assert.match(workflow, /npm run quality:native/);
  assert.doesNotMatch(workflow, /pull_request|upload-artifact|qualify:deck/);
});

test("quality orchestrator records the privacy boundary and separates portable from native", () => {
  const script = read("scripts/run-quality-pipeline.mjs");
  assert.match(script, /clientContentIncluded:\s*false/);
  assert.match(script, /--native-canary/);
  assert.match(script, /--ci/);
  assert.match(script, /synthetic precision-layout canary/);
  assert.match(script, /Client presentations and qualification images are never uploaded or archived/);
});
