import assert from "node:assert/strict";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { bindNativeMeasurement, calculateNativeCellClearances, calculateNativeTextOverflowEdges, remapSingleSlideNativeMeasurement } from "../src/lib/native-measurement";
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
    slides: Array<{ number: number; shapes: Array<{ nativeShapeId?: string; name?: string; text?: { renderedBoundsPt?: { left: number; top: number; width: number; height: number } }; table?: { rowHeightsPt: number[]; columnWidthsPt: number[]; cells: Array<{ renderedTextBoundsPt?: { left: number; top: number; width: number; height: number }; marginsPt?: { left: number } }> } }> }>;
  };
  nativeMeasurementCapabilities(platform?: string): { available: boolean; adapter: string };
};

test("native measurement automation closes its exact temporary presentation after success or failure", () => {
  assert.doesNotMatch(POWERPOINT_MEASUREMENT_SCRIPT, /set targetPresentation to active presentation/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /full name of candidatePresentation as text/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /count of slides of targetPresentation/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /on error measurementError number measurementErrorNumber/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /close targetPresentation saving no/i);
  assert.match(POWERPOINT_MEASUREMENT_SCRIPT, /if \(full name of candidatePresentation as text\) is sourcePath then\s+close candidatePresentation saving no/i);
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
    "CELL\t1\t2\t1\t2\t351.84\t120\t300\t44\t5\t5\t4\t4\t305\t0\t60\t16\t8\t1\tanchor top",
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
  assert.ok(Math.abs((parsed.slides[0].shapes[1].table?.cells[1].renderedTextBoundsPt?.left ?? 0) - 5) < 0.001);
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

test("isolated native measurements remap PowerPoint slide one to the original slide number", () => {
  const native: NativeMeasurementResult = {
    status: "ready",
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    slideCount: 1,
    slides: [{
      number: 1,
      shapeCount: 1,
      shapes: [{ slideNumber: 1, shapeIndex: 1, nativeShapeId: "7", zOrder: 1, boundsPt: { left: 10, top: 20, width: 30, height: 40 }, rotation: 0, hasTextFrame: false, hasTable: false }],
    }],
    warnings: [],
  };
  const remapped = remapSingleSlideNativeMeasurement(native, 13);
  assert.equal(remapped.slideCount, 1);
  assert.equal(remapped.slides[0].number, 13);
  assert.equal(remapped.slides[0].shapes[0].slideNumber, 13);
});

test("native measurement binding can return only one requested source slide", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-slide-binding-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const audit = await auditPptx(new Uint8Array(await fs.readFile(filePath)));
  const deck: DeckJob = { id: "native-slide-binding", name: "synthetic.pptx", sourceResourceId: "native-slide-binding-source", sourceSha256: "d".repeat(64), operationScope: "reflow", templateClassification: audit.classification, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  const object = deck.scene.objects.find((candidate) => candidate.slideNumber === 2);
  assert.ok(object);
  const native: NativeMeasurementResult = {
    status: "ready",
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    sourceSha256: deck.sourceSha256,
    slideCount: 1,
    slides: [{ number: 2, shapeCount: 1, shapes: [
      { slideNumber: 2, shapeIndex: 1, nativeShapeId: object.shapeId, name: object.name, zOrder: object.zIndex + 1, boundsPt: { left: 100, top: 100, width: 200, height: 50 }, rotation: 0, hasTextFrame: true, hasTable: false },
    ] }],
    warnings: [],
  };
  const packet = bindNativeMeasurement(deck, native, { slideNumbers: [2] });
  assert.ok(packet.objects.length > 0);
  assert.ok(packet.objects.every((candidate) => candidate.slideNumber === 2));
  assert.equal(packet.objects.find((candidate) => candidate.objectId === object.id)?.provenance.authority, "powerpoint-native");
});

test("native measurement binding ignores duplicate legacy scene IDs instead of consuming another PowerPoint shape", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-duplicate-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const audit = await auditPptx(new Uint8Array(await fs.readFile(filePath)));
  const deck: DeckJob = { id: "native-duplicate", name: "synthetic.pptx", sourceResourceId: "native-duplicate-source", sourceSha256: "c".repeat(64), operationScope: "reflow", templateClassification: audit.classification, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  const object = deck.scene.objects[0];
  assert.ok(object);
  deck.scene = { ...deck.scene, objects: [object, { ...object }, ...deck.scene.objects.slice(1)] };
  const native: NativeMeasurementResult = {
    status: "ready",
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    sourceSha256: deck.sourceSha256,
    slides: [{ number: object.slideNumber, shapeCount: 2, shapes: [
      { slideNumber: object.slideNumber, shapeIndex: 1, nativeShapeId: object.shapeId, name: object.name, zOrder: 1, boundsPt: { left: 20, top: 20, width: 200, height: 40 }, rotation: 0, hasTextFrame: true, hasTable: false },
      { slideNumber: object.slideNumber, shapeIndex: 2, nativeShapeId: "other", name: "Other", zOrder: 2, boundsPt: { left: 400, top: 300, width: 100, height: 30 }, rotation: 0, hasTextFrame: true, hasTable: false },
    ] }],
    warnings: [],
  };
  const packet = bindNativeMeasurement(deck, native);
  assert.equal(packet.objects.filter((candidate) => candidate.objectId === object.id).length, 1);
  assert.equal(packet.objects.find((candidate) => candidate.objectId === object.id)?.measuredGeometryPt?.left, 20);
  assert.ok(packet.warnings.some((warning) => warning.includes(`Duplicate scene object ${object.id}`)));
});

test("native cell clearance retains negative clipping evidence instead of clamping it away", () => {
  const clearances = calculateNativeCellClearances({ row: 1, column: 1, boundsPt: { left: 0, top: 0, width: 100, height: 20 }, marginsPt: { left: 4, right: 4, top: 4, bottom: 4 }, renderedTextBoundsPt: { left: -2, top: 0, width: 104, height: 16 }, textCoordinateSpace: "cell-relative", textLength: 10, lineCount: 1, verticalAnchor: "anchor top" });
  assert.deepEqual(clearances, { left: -2, right: -2, top: 4, bottom: 0 });
});

test("native text overflow tolerates inset and glyph overhang but detects escape from the outer frame", () => {
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
  assert.deepEqual(calculateNativeTextOverflowEdges(object), []);
  object.text.renderedBoundsPt.width = 98;
  assert.deepEqual(calculateNativeTextOverflowEdges(object), ["right"]);
});

test("native measurement capabilities keep platform adapters explicit", () => {
  assert.equal(nativeMeasurementCapabilities("win32").adapter, "windows-powerpoint-com-pending");
  assert.equal(nativeMeasurementCapabilities("linux").adapter, "ooxml-fallback");
});
