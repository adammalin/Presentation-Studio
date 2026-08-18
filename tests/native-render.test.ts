import assert from "node:assert/strict";
import { chmodSync, mkdtempSync, mkdirSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import test from "node:test";

const require = createRequire(import.meta.url);
const {
  POWERPOINT_RENDER_SCRIPT,
  bundledPdfRasterizerAvailable,
  classifyPowerPointAutomationError,
  jpegDimensions,
  pngDimensions,
  rasterizePdfWithPdfJs,
  slideNumberFromFile,
  validatePdfRasterizer,
} = require("../electron/native-render.cjs") as {
  POWERPOINT_RENDER_SCRIPT: string;
  bundledPdfRasterizerAvailable(): boolean;
  classifyPowerPointAutomationError(error: unknown): string;
  jpegDimensions(bytes: Uint8Array): { width: number; height: number };
  pngDimensions(bytes: Uint8Array): { width: number; height: number };
  rasterizePdfWithPdfJs(input: { pdfPath: string; outputDirectory: string; width: number; format: "jpeg" | "png" }): Promise<void>;
  slideNumberFromFile(fileName: string): number;
  validatePdfRasterizer(candidate?: string): string | null;
};

function minimalPdf(): Buffer {
  const content = "BT /F1 18 Tf 20 50 Td (Presentation Studio) Tj ET\n";
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}endstream`,
  ];
  let source = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(source));
    source += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xref = Buffer.byteLength(source);
  source += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  source += offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("");
  source += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(source);
}

test("native rasterizer readiness resolves symlinks and proves the executable can start", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "presentation studio rasterizer "));
  try {
    const binDirectory = path.join(root, "bundled poppler", "bin");
    mkdirSync(binDirectory, { recursive: true });
    const target = path.join(binDirectory, process.platform === "win32" ? "pdftoppm.exe" : "pdftoppm");
    if (process.platform === "win32") symlinkSync(process.execPath, target);
    else {
      writeFileSync(target, "#!/usr/bin/env bash\n[ \"$1\" = \"-v\" ]\n", { mode: 0o755 });
      chmodSync(target, 0o755);
    }
    const linkedDirectory = path.join(root, "consumer project", "node_modules", ".bin");
    mkdirSync(linkedDirectory, { recursive: true });
    const link = path.join(linkedDirectory, process.platform === "win32" ? "pdftoppm.exe" : "pdftoppm");
    symlinkSync(target, link);

    assert.equal(validatePdfRasterizer(link), realpathSync(target));

    const broken = path.join(binDirectory, process.platform === "win32" ? "broken-pdftoppm.exe" : "broken-pdftoppm");
    writeFileSync(broken, process.platform === "win32" ? "not an executable" : "#!/usr/bin/env bash\nexit 1\n", { mode: 0o755 });
    chmodSync(broken, 0o755);
    assert.equal(validatePdfRasterizer(broken), null);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("bundled PDF.js rasterizer produces slide images without a system pdftoppm", async () => {
  assert.equal(bundledPdfRasterizerAvailable(), true);
  const root = mkdtempSync(path.join(os.tmpdir(), "presentation studio bundled rasterizer "));
  try {
    const pdfPath = path.join(root, "synthetic.pdf");
    writeFileSync(pdfPath, minimalPdf());
    await rasterizePdfWithPdfJs({ pdfPath, outputDirectory: root, width: 800, format: "png" });
    const data = readFileSync(path.join(root, "slide-1.png"));
    assert.deepEqual(pngDimensions(data), { width: 800, height: 400 });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("native render automation closes its exact temporary presentation after success or failure", () => {
  assert.doesNotMatch(POWERPOINT_RENDER_SCRIPT, /set targetPresentation to active presentation/i);
  assert.match(POWERPOINT_RENDER_SCRIPT, /full name of candidatePresentation as text/i);
  assert.match(POWERPOINT_RENDER_SCRIPT, /save targetPresentation/i);
  assert.match(POWERPOINT_RENDER_SCRIPT, /on error renderError number renderErrorNumber/i);
  assert.match(POWERPOINT_RENDER_SCRIPT, /close targetPresentation saving no/i);
  assert.match(POWERPOINT_RENDER_SCRIPT, /if \(full name of candidatePresentation as text\) is sourcePath then\s+close candidatePresentation saving no/i);
});

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

test("PowerPoint automation diagnostics surface actionable conditions without echoing the AppleScript command", () => {
  const { describePowerPointAutomationError, macSessionLocked } = require("../electron/powerpoint-automation-error.cjs") as { describePowerPointAutomationError(error: unknown, action?: string): { status: string; reason: string; message: string }; macSessionLocked(output: string): boolean };
  const locked = describePowerPointAutomationError({ stderr: "The Mac is locked and automatic unlock could not unlock it." }, "rendering");
  assert.equal(locked.reason, "mac-session-locked");
  assert.match(locked.message, /unlock the Mac/i);
  const timeout = describePowerPointAutomationError({ killed: true, signal: "SIGTERM", message: "Command failed: /usr/bin/osascript -e very long script" }, "measurement");
  assert.equal(timeout.reason, "powerpoint-automation-timeout");
  assert.doesNotMatch(timeout.message, /very long script/i);
  const failed = describePowerPointAutomationError({ message: "Command failed: /usr/bin/osascript -e long script\nMicrosoft PowerPoint returned error -42" }, "rendering");
  assert.equal(failed.reason, "powerpoint-automation-failed");
  assert.match(failed.message, /error -42/i);
  assert.doesNotMatch(failed.message, /osascript/i);
  assert.equal(macSessionLocked('"CGSSessionScreenIsLocked"=Yes'), true);
  assert.equal(macSessionLocked('"CGSSessionScreenIsLocked"=No'), false);
});

test("PowerPoint startup-window recovery is bounded to an empty presentation session", async () => {
  const { POWERPOINT_CLOSE_EXACT_SCRIPT, runPowerPointAutomationWithStartupRecovery } = require("../electron/powerpoint-automation-error.cjs") as {
    POWERPOINT_CLOSE_EXACT_SCRIPT: string;
    runPowerPointAutomationWithStartupRecovery(options: { action: string; run(): Promise<string>; presentationCount(): Promise<number>; quit(): Promise<void> }): Promise<string>;
  };
  assert.doesNotMatch(POWERPOINT_CLOSE_EXACT_SCRIPT, /active presentation/i);
  assert.match(POWERPOINT_CLOSE_EXACT_SCRIPT, /application "Microsoft PowerPoint" is running/i);
  assert.match(POWERPOINT_CLOSE_EXACT_SCRIPT, /full name of candidatePresentation as text/i);
  assert.match(POWERPOINT_CLOSE_EXACT_SCRIPT, /close candidatePresentation saving no/i);
  let attempts = 0;
  let quits = 0;
  const recovered = await runPowerPointAutomationWithStartupRecovery({
    action: "rendering",
    run: async () => {
      attempts += 1;
      if (attempts === 1) throw { stderr: "Microsoft PowerPoint got an error: An error of type -9074 has occurred. (-9074)" };
      return "rendered";
    },
    presentationCount: async () => 0,
    quit: async () => { quits += 1; },
  });
  assert.equal(recovered, "rendered");
  assert.equal(attempts, 2);
  assert.equal(quits, 1);

  attempts = 0;
  quits = 0;
  await assert.rejects(() => runPowerPointAutomationWithStartupRecovery({
    action: "measurement",
    run: async () => {
      attempts += 1;
      throw { stderr: "An error of type -9074 has occurred. (-9074)" };
    },
    presentationCount: async () => 1,
    quit: async () => { quits += 1; },
  }));
  assert.equal(attempts, 1);
  assert.equal(quits, 0);
});
