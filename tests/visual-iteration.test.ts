import assert from "node:assert/strict";
import test from "node:test";
import { decideVisualIteration } from "../src/lib/visual-iteration";

const hashA = "a".repeat(64);
const hashB = "b".repeat(64);
const base = { rationale: "The hierarchy is clearer and the grouping feels intentional.", slideNumber: 2, inspectionRevision: "revision-2", currentRasterSha256: hashA, proposalRasterSha256: hashB, changedPixelRatio: .12, improvements: ["optical left-edge deviation"], regressions: [] };

test("visual loop records a better verdict only when native pixels changed without metric regression", () => {
  const result = decideVisualIteration({ ...base, priorHistory: [], requestedVerdict: "better" });
  assert.equal(result.verdict, "better");
  assert.equal(result.entry.attempt, 1);
  assert.equal(result.rejected, false);
});

test("visual loop withholds a claimed improvement when objective metrics regress", () => {
  const result = decideVisualIteration({ ...base, priorHistory: [], requestedVerdict: "better", regressions: ["text overflow"] });
  assert.equal(result.verdict, "revise");
  assert.match(result.entry.rationale, /withheld/i);
  assert.deepEqual(result.entry.metrics.regressions, ["text overflow"]);
});

test("the third unresolved AI revision is rejected instead of looping forever", () => {
  const first = decideVisualIteration({ ...base, priorHistory: [], requestedVerdict: "revise" }).entry;
  const second = decideVisualIteration({ ...base, priorHistory: [first], requestedVerdict: "revise" }).entry;
  const third = decideVisualIteration({ ...base, priorHistory: [first, second], requestedVerdict: "revise" });
  assert.equal(third.verdict, "reject");
  assert.equal(third.entry.attempt, 3);
  assert.match(third.entry.rationale, /exhausted/i);
});
