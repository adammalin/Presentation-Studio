import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";

const require = createRequire(import.meta.url);
const {
  classifyPowerPointAutomationError,
  jpegDimensions,
  slideNumberFromFile,
} = require("../electron/native-render.cjs") as {
  classifyPowerPointAutomationError(error: unknown): string;
  jpegDimensions(bytes: Uint8Array): { width: number; height: number };
  slideNumberFromFile(fileName: string): number;
};

test("reads native JPEG dimensions before slide images enter the renderer", () => {
  const onePixelJpeg = Buffer.from("/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAf/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIAAwAAABAf/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPxB//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPxB//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxB//9k=", "base64");
  assert.deepEqual(jpegDimensions(onePixelJpeg), { width: 1, height: 1 });
  assert.throws(() => jpegDimensions(Uint8Array.of(0, 1, 2, 3)), /invalid JPEG/i);
});

test("sort keys and permission failures are classified deterministically", () => {
  assert.equal(slideNumberFromFile("slide-21.jpg"), 21);
  assert.ok(Number.isNaN(slideNumberFromFile("cover.jpg")));
  assert.equal(classifyPowerPointAutomationError({ stderr: "Not authorized to send Apple events. (-1743)" }), "permission-required");
  assert.equal(classifyPowerPointAutomationError({ message: "PowerPoint crashed" }), "failed");
});
