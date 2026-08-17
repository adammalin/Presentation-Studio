import assert from "node:assert/strict";
import test from "node:test";
import { nativeTextOverflows } from "../src/lib/fresh-composition-qa";
import type { NativeMeasurementResult, NativeShapeMeasurement } from "../src/lib/desktop";

function shape(name: string, textLength: number, renderedWidth: number): NativeShapeMeasurement {
  return {
    slideNumber: 1,
    shapeIndex: 1,
    name,
    zOrder: 1,
    boundsPt: { left: 10, top: 10, width: 100, height: 30 },
    rotation: 0,
    hasTextFrame: true,
    hasTable: false,
    text: {
      marginsPt: { left: 0, right: 0, top: 0, bottom: 0 },
      renderedBoundsPt: { left: 10, top: 10, width: renderedWidth, height: 20 },
      coordinateSpace: "slide",
      textLength,
      lineCount: textLength ? 1 : 0,
      verticalAnchor: "top",
    },
  };
}

function measurement(shapes: NativeShapeMeasurement[]): NativeMeasurementResult {
  return { status: "ready", adapter: "macos-powerpoint-applescript", authority: "powerpoint-native", slides: [{ number: 1, shapeCount: shapes.length, shapes }], warnings: [] };
}

test("fresh-composition QA ignores empty decoration frames and normal glyph overhang", () => {
  assert.deepEqual(nativeTextOverflows(measurement([shape("empty rule", 0, 125), shape("PowerPoint empty rule", 1, 125), shape("right aligned footer", 10, 102)])), []);
});

test("fresh-composition QA reports material PowerPoint-native text overflow", () => {
  assert.deepEqual(nativeTextOverflows(measurement([shape("body", 40, 103)])), [{ slideNumber: 1, name: "body", edges: ["right"] }]);
});
