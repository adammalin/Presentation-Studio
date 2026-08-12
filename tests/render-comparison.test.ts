import assert from "node:assert/strict";
import test from "node:test";
import { compareRgbaPixels } from "../src/lib/render-comparison";

test("pixel comparison reports exact matches and bounded material changes", () => {
  const current = new Uint8ClampedArray([
    255, 255, 255, 255, 255, 255, 255, 255,
    255, 255, 255, 255, 255, 255, 255, 255,
  ]);
  assert.equal(compareRgbaPixels(current, current.slice(), 2, 2).exactPixelMatch, true);
  const proposal = current.slice();
  proposal[4] = 0;
  proposal[5] = 102;
  proposal[6] = 44;
  const metrics = compareRgbaPixels(current, proposal, 2, 2);
  assert.equal(metrics.exactPixelMatch, false);
  assert.equal(metrics.changedPixelCount, 1);
  assert.equal(metrics.changedPixelRatio, .25);
  assert.deepEqual(metrics.changedBounds?.normalized, { x: .5, y: 0, width: .5, height: .5 });
  assert.equal(metrics.maximumChannelDelta, 255);
});
