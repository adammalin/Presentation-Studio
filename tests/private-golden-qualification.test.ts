import assert from "node:assert/strict";
import test from "node:test";
import { parsePrivateGoldenManifest } from "../scripts/qualify-private-golden";

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
