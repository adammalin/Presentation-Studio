import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { evaluateBlackBoxAgentRun, parseBlackBoxAgentResult, prepareBlackBoxAgentRun } from "../scripts/black-box-agent-acceptance";
import { parsePrivateGoldenManifest, privateGoldenContentCoverage } from "../scripts/qualify-private-golden";

const base = {
  schema: "presentation-studio/private-golden-manifest",
  version: 1,
  id: "private-golden-test",
  source: { path: "/private/source.pptx", sha256: "a".repeat(64) },
  benchmark: { path: "/private/benchmark.pptx", sha256: "b".repeat(64) },
  cases: [{ id: "representative-layout", sourceSlide: 2, benchmarkSlides: [2], communicationJob: "Clarify one source-bound assertion and its evidence.", reviewFocus: ["hierarchy", "layout-balance", "editability", "source-intent"] }],
};

test("private golden manifests remain hash-pinned and select bounded representative cases", () => {
  const manifest = parsePrivateGoldenManifest(base);
  assert.equal(manifest.source.sha256, "a".repeat(64));
  assert.equal(manifest.benchmark.sha256, "b".repeat(64));
  assert.equal(manifest.cases[0].sourceSlide, 2);
  assert.deepEqual(manifest.cases[0].reviewFocus, ["hierarchy", "layout-balance", "editability", "source-intent"]);
});

test("version 2 private golden manifests pin the authorized template and design mode", () => {
  const manifest = parsePrivateGoldenManifest({ ...base, version: 2, template: { path: "/private/authorized-template.potx", sha256: "c".repeat(64) }, designMode: "shared" });
  assert.equal(manifest.version, 2);
  assert.equal(manifest.template?.sha256, "c".repeat(64));
  assert.equal(manifest.designMode, "shared");
});

test("private golden manifests reject unpinned, duplicate, or unbounded cases", () => {
  assert.throws(() => parsePrivateGoldenManifest({ ...base, source: { ...base.source, sha256: "not-a-hash" } }), /SHA-256/i);
  assert.throws(() => parsePrivateGoldenManifest({ ...base, cases: [...base.cases, { ...base.cases[0], id: "duplicate-slide" }] }), /unique positive source slide/i);
  assert.throws(() => parsePrivateGoldenManifest({ ...base, cases: [] }), /1–12/i);
  assert.throws(() => parsePrivateGoldenManifest({ ...base, version: 2, designMode: "template" }), /requires an authorized template/i);
});

test("private golden content coverage detects ordinal drift in a human-edited benchmark", () => {
  const source = "Workshop Philip Boudreaux attended the Air Barrier Association conference in Minneapolis Q3";
  const correct = "Philip Boudreaux attended the Air Barrier Association conference in Minneapolis during Q3";
  const shifted = "Hot Water and Hot Air Forums was held in Phoenix during Q2";
  assert.ok(privateGoldenContentCoverage(source, [correct]) >= 0.8);
  assert.ok(privateGoldenContentCoverage(source, [shifted]) < 0.3);
});

test("black-box acceptance prepares a benchmark-blind prompt and scores a structured fresh-agent result", async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-black-box-"));
  const qualificationPath = path.join(root, "qualification.json");
  await fs.writeFile(qualificationPath, JSON.stringify({
    schema: "presentation-studio/private-golden-qualification",
    version: 2,
    manifest: { id: "synthetic-golden" },
    source: { sha256: "a".repeat(64) },
    cases: [{ id: "metric-grid", sourceSlide: 2, communicationJob: "Compose the exact metrics as one coherent system.", reviewFocus: ["hierarchy", "source-intent"] }],
  }));
  const prepared = await prepareBlackBoxAgentRun(qualificationPath, root);
  const prompt = await fs.readFile(prepared.promptPath, "utf8");
  assert.match(prompt, /context-free product acceptance agent/i);
  assert.doesNotMatch(prompt, /human-cleaned.*path|benchmark\.pptx/i);
  const result = {
    schema: "presentation-studio/black-box-agent-result",
    version: 1,
    runId: prepared.run.id,
    completedAt: "2026-08-20T18:00:00.000Z",
    designStandardVersion: prepared.run.designStandardVersion,
    sourceSha256: prepared.run.sourceSha256,
    noSaveOrExport: true,
    cases: [{ id: "metric-grid", sourceSlide: 2, status: "pass", attempts: 2, summary: "The exact native candidate is stronger.", defects: [] }],
  };
  assert.equal(parseBlackBoxAgentResult(result, prepared.run).cases[0].status, "pass");
  await fs.writeFile(prepared.resultPath, JSON.stringify(result));
  const evaluated = await evaluateBlackBoxAgentRun(prepared.runPath, prepared.resultPath, root);
  assert.equal(evaluated.summary.result.passCount, 1);
  assert.equal(evaluated.summary.trend, "baseline");
});
