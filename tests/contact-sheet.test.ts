import assert from "node:assert/strict";
import test from "node:test";
import { CONTACT_SHEET_PAGE_SIZE, planContactSheet } from "../src/lib/contact-sheet";
import type { NativeSlideRender } from "../src/lib/desktop";

function slides(count: number): NativeSlideRender[] {
  return Array.from({ length: count }, (_value, index) => ({
    number: index + 1,
    mimeType: "image/jpeg",
    width: 1600,
    height: 900,
    sha256: String(index + 1).padStart(64, "0"),
    bytes: new Uint8Array([index + 1]),
  }));
}

test("deck contact sheets page a 200-slide review into bounded PowerPoint-native overviews", () => {
  const first = planContactSheet(slides(200), 1);
  const last = planContactSheet(slides(200), 5);
  assert.equal(CONTACT_SHEET_PAGE_SIZE, 40);
  assert.equal(first.pageCount, 5);
  assert.equal(first.placements.length, 40);
  assert.equal(first.firstSlideNumber, 1);
  assert.equal(first.lastSlideNumber, 40);
  assert.equal(last.firstSlideNumber, 161);
  assert.equal(last.lastSlideNumber, 200);
  assert.ok(first.width >= 1600 && first.height >= 1600);
  assert.equal(new Set(first.placements.map((placement) => `${placement.x}:${placement.y}`)).size, 40);
});

test("contact-sheet planning rejects stale or unbounded page requests", () => {
  assert.throws(() => planContactSheet(slides(41), 3), /page 1 to 2/i);
  assert.throws(() => planContactSheet(slides(1), 1, 51), /page size/i);
  assert.throws(() => planContactSheet([], 1), /at least one/i);
});
