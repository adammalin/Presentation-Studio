import assert from "node:assert/strict";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { bindNativeMeasurement, calculateNativeCellClearances, calculateNativeTextOverflowEdges } from "../src/lib/native-measurement";
import { auditPptx } from "../src/lib/pptx-audit";
import { compilePresentationScene } from "../src/lib/scene-graph";
import type { NativeMeasurementResult } from "../src/lib/desktop";
import type { DeckJob } from "../src/types";

const require = createRequire(import.meta.url);
const { POWERPOINT_MEASUREMENT_SCRIPT, parsePowerPointMeasurement, nativeMeasurementCapabilities } = require("../electron/native-measurement.cjs") as {
  POWERPOINT_MEASUREMENT_SCRIPT: string;
  parsePowerPointMeasurement(text: string, options?: { sourceSha256?: string }): {
    status: string;
    authority: string;
    slideCount: number;
    slides: Array<{ number: number; shapes: Array<{ nativeShapeId?: string; name?: string; text?: { renderedBoundsPt?: { left: number; top: number; width: number; height: number } }; table?: { rowHeightsPt: number[]; columnWidthsPt: number[]; cells: Array<{ renderedTextBoundsPt?: { width: number }; marginsPt?: { left: number } }> } }> }>;
  };
  nativeMeasurementCapabilities(platform?: string): { available: boolean; adapter: string };
};

test("native measurement automation closes its exact temporary presentation after success or failure", () => {
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /set targetPresentation to active presentation/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /count of slides of targetPresentation/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /on error measurementError number measurementErrorNumber/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /close targetPresentation saving no/i);
});

test("native PowerPoint measurement parser retains rendered text and cell geometry without copy", () => {
  const parsed = parsePowerPointMeasurement([
    "DECK\t16.111.2\t1",
    "SLIDE\t1\t2",
    "SHAPE\t1\t1\t1\t51.84\t34.56\t500\t40\t0\ttrue\tfalse\t6\tTitle 6",
    "TEXT\t1\t1\t0\t0\t0\t0\t51.84\t38.2\t420\t28.8\t37\t1\tanchor middle",
    "SHAPE\t1\t2\t2\t51.84\t120\t600\t120\t0\tfalse\ttrue",
    "TABLE\t1\t2\t1\t2",
    "ROW\t1\t2\t1\t44",
    "COL\t1\t2\t1\t300",
    "COL\t1\t2\t2\t300",
    "CELL\t1\t2\t1\t1\t51.84\t120\t300\t44\t5\t5\t4\t4\t5\t0\t80\t16\t10\t1\tanchor top",
    "CELL\t1\t2\t1\t2\t351.84\t120\t300\t44\t5\t5\t4\t4\t5\t0\t60\t16\t8\t1\tanchor top",
  ].join("\n"), { sourceSha256: "a".repeat(64) });
  assert.equal(parsed.status, "ready");
  assert.equal(parsed.authority, "powerpoint-native");
  assert.equal(parsed.slideCount, 1);
  assert.equal(parsed.slides[0].shapes[0].text?.renderedBoundsPt?.left, 51.84);
  assert.equal(parsed.slides[0].shapes[0].nativeShapeId, "6");
  assert.equal(parsed.slides[0].shapes[0].name, "Title 6");
  assert.deepEqual(parsed.slides[0].shapes[1].table?.rowHeightsPt, [44]);
  assert.deepEqual(parsed.slides[0].shapes[1].table?.columnWidthsPt, [300, 300]);
  assert.equal(parsed.slides[0].shapes[1].table?.cells[0].marginsPt?.left, 5);
  assert.equal(parsed.slides[0].shapes[1].table?.cells[0].renderedTextBoundsPt?.width, 80);
  assert.doesNotMatch(JSON.stringify(parsed), /technical copy/i);
});

test("native measurements bind by stable PowerPoint shape ID before z-order or proposal geometry", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-binding-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const audit = await auditPptx(new Uint8Array(await fs.readFile(filePath)));
  const deck: DeckJob = { id: "native-binding", name: "synthetic.pptx", sourceResourceId: "native-binding-source", sourceSha256: "b".repeat(64), operationScope: "reflow", templateClassification: audit.classification, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  const object = deck.scene.objects.find((candidate) => candidate.sourceElement === "p:sp");
  assert.ok(object);
  const native: NativeMeasurementResult = {
    status: "ready",
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    sourceSha256: deck.sourceSha256,
    slides: [{
      number: object.slideNumber,
      shapeCount: 2,
      shapes: [
        { slideNumber: object.slideNumber, shapeIndex: 1, nativeShapeId: "9999", name: "AI decoration", zOrder: object.zIndex + 1, boundsPt: { left: 1, top: 1, width: 5, height: 5 }, rotation: 0, hasTextFrame: false, hasTable: false },
        { slideNumber: object.slideNumber, shapeIndex: 2, nativeShapeId: object.shapeId, name: object.name, zOrder: 99, boundsPt: { left: 300, top: 200, width: 250, height: 40 }, rotation: 0, hasTextFrame: true, hasTable: false },
      ],
    }],
    warnings: [],
  };
  const packet = bindNativeMeasurement(deck, native);
  const bound = packet.objects.find((candidate) => candidate.objectId === object.id);
  assert.equal(bound?.binding.method, "shape-id");
  assert.equal(bound?.binding.nativeShapeIndex, 2);
  assert.equal(bound?.measuredGeometryPt?.left, 300);
});

test("native cell clearance retains negative clipping evidence instead of clamping it away", () => {
  const clearances = calculateNativeCellClearances({ row: 1, column: 1, boundsPt: { left: 0, top: 0, width: 100, height: 20 }, marginsPt: { left: 4, right: 4, top: 4, bottom: 4 }, renderedTextBoundsPt: { left: -2, top: 0, width: 104, height: 16 }, textCoordinateSpace: "cell-relative", textLength: 10, lineCount: 1, verticalAnchor: "anchor top" });
  assert.deepEqual(clearances, { left: -2, right: -2, top: 4, bottom: 0 });
});

test("native text overflow uses the inset text region rather than the outer shape edge", () => {
  const object = {
    objectId: "text",
    shapeId: "1",
    slideNumber: 1,
    sourceGeometryPt: { left: 10, top: 10, width: 100, height: 30 },
    measuredGeometryPt: { left: 10, top: 10, width: 100, height: 30 },
    text: { marginsPt: { left: 6, right: 6, top: 3, bottom: 3 }, renderedBoundsPt: { left: 15, top: 13, width: 80, height: 20 }, coordinateSpace: "slide" as const, textLength: 10, lineCount: 1, verticalAnchor: "anchor top" },
    binding: { method: "shape-id" as const, confidence: "high" as const },
    provenance: { authority: "powerpoint-native" as const, adapter: "macos-powerpoint-applescript", confidence: "high" as const, note: "test" },
  };
  assert.deepEqual(calculateNativeTextOverflowEdges(object), ["left"]);
});

test("native measurement capabilities keep platform adapters explicit", () => {
  assert.equal(nativeMeasurementCapabilities("win32").adapter, "windows-powerpoint-com-pending");
  assert.equal(nativeMeasurementCapabilities("linux").adapter, "ooxml-fallback");
});
