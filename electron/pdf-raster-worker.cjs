const fsSync = require("node:fs");
const fs = require("node:fs/promises");
const path = require("node:path");

const MAX_NATIVE_SLIDES = 1_000;

function validatedRequest(raw) {
  const request = JSON.parse(raw || "{}");
  const pdfPath = path.resolve(String(request.pdfPath || ""));
  const outputDirectory = path.resolve(String(request.outputDirectory || ""));
  const width = Number(request.width);
  const format = request.format === "png" ? "png" : request.format === "jpeg" ? "jpeg" : undefined;
  if (!path.isAbsolute(String(request.pdfPath || "")) || !fsSync.statSync(pdfPath).isFile()) throw new Error("The isolated PDF worker requires an existing absolute PDF path.");
  if (!path.isAbsolute(String(request.outputDirectory || "")) || !fsSync.statSync(outputDirectory).isDirectory()) throw new Error("The isolated PDF worker requires an existing absolute output folder.");
  if (!Number.isInteger(width) || width < 800 || width > 3_000) throw new Error("The isolated PDF worker width must be from 800 to 3,000 pixels.");
  if (!format) throw new Error("The isolated PDF worker format must be JPEG or PNG.");
  return { pdfPath, outputDirectory, width, format };
}

async function rasterize(request) {
  // Keep the native Skia dependency entirely inside this disposable process.
  // A malformed or unusually complex PDF must never be able to terminate the
  // Electron main process that owns the user's open project.
  const canvasApi = require("@napi-rs/canvas");
  for (const globalName of ["DOMMatrix", "ImageData", "Path2D"]) {
    if (!globalThis[globalName] && canvasApi[globalName]) globalThis[globalName] = canvasApi[globalName];
  }
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(await fs.readFile(request.pdfPath)),
    disableWorker: true,
    isEvalSupported: false,
    useSystemFonts: true,
  });
  const document = await loadingTask.promise;
  try {
    if (document.numPages < 1) throw new Error("The PowerPoint PDF render does not contain any pages.");
    if (document.numPages > MAX_NATIVE_SLIDES) throw new Error("The presentation exceeds the 1,000-slide native render limit.");
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await document.getPage(pageNumber);
      try {
        const baseViewport = page.getViewport({ scale: 1 });
        const viewport = page.getViewport({ scale: request.width / baseViewport.width });
        const canvas = canvasApi.createCanvas(Math.round(viewport.width), Math.round(viewport.height));
        const context = canvas.getContext("2d");
        await page.render({ canvas, canvasContext: context, viewport, intent: "print", background: "rgb(255,255,255)" }).promise;
        const extension = request.format === "png" ? "png" : "jpg";
        // @napi-rs/canvas.encode() uses a native asynchronous worker. Electron
        // crash evidence showed that path dereferencing a detached ArrayBuffer.
        // Synchronous toBuffer() keeps ownership deterministic inside this
        // already-isolated process.
        const encoded = request.format === "png" ? canvas.toBuffer("image/png") : canvas.toBuffer("image/jpeg", 90);
        await fs.writeFile(path.join(request.outputDirectory, `slide-${pageNumber}.${extension}`), encoded, { mode: 0o600 });
      } finally {
        page.cleanup();
      }
    }
  } finally {
    await document.destroy();
  }
}

void rasterize(validatedRequest(process.argv[2])).catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 1;
});
