import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import test from "node:test";

const require = createRequire(import.meta.url);
const { validateRasterRequest } = require("../electron/chromium-pdf-raster.cjs") as {
  validateRasterRequest(input: Record<string, unknown>): { pdfPath: string; outputDirectory: string; width: number; format: "jpeg" | "png" };
};

test("Chromium PDF raster requests are restricted to existing local paths and bounded output", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "presentation studio chromium raster "));
  try {
    const pdfPath = path.join(root, "source.pdf");
    const outputDirectory = path.join(root, "output");
    writeFileSync(pdfPath, "%PDF-1.4\n");
    mkdirSync(outputDirectory);
    assert.deepEqual(validateRasterRequest({ pdfPath, outputDirectory, width: 1400, format: "png" }), { pdfPath, outputDirectory, width: 1400, format: "png" });
    assert.throws(() => validateRasterRequest({ pdfPath: "relative.pdf", outputDirectory, width: 1400, format: "png" }), /absolute PDF path/i);
    assert.throws(() => validateRasterRequest({ pdfPath, outputDirectory, width: 799, format: "png" }), /800 to 3,000/i);
    assert.throws(() => validateRasterRequest({ pdfPath, outputDirectory, width: 1400, format: "gif" }), /JPEG or PNG/i);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
