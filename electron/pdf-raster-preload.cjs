const { contextBridge } = require("electron");
const fsSync = require("node:fs");
const fs = require("node:fs/promises");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const MAX_NATIVE_SLIDES = 1_000;

function validateRequest(input) {
  const request = {
    pdfPath: path.resolve(String(input?.pdfPath || "")),
    outputDirectory: path.resolve(String(input?.outputDirectory || "")),
    width: Number(input?.width),
    format: input?.format === "png" ? "png" : input?.format === "jpeg" ? "jpeg" : undefined,
  };
  if (!path.isAbsolute(String(input?.pdfPath || "")) || !fsSync.statSync(request.pdfPath).isFile()) throw new Error("The Chromium PDF renderer requires an existing absolute PDF path.");
  if (!path.isAbsolute(String(input?.outputDirectory || "")) || !fsSync.statSync(request.outputDirectory).isDirectory()) throw new Error("The Chromium PDF renderer requires an existing absolute output folder.");
  if (!Number.isInteger(request.width) || request.width < 800 || request.width > 3_000) throw new Error("The Chromium PDF renderer width must be from 800 to 3,000 pixels.");
  if (!request.format) throw new Error("The Chromium PDF renderer format must be JPEG or PNG.");
  return request;
}

function canvasBytes(canvas, format) {
  return new Promise((resolve, reject) => {
    const mimeType = format === "png" ? "image/png" : "image/jpeg";
    canvas.toBlob(async (blob) => {
      if (!blob) return reject(new Error("Chromium could not encode a PDF page image."));
      try {
        resolve(Buffer.from(await blob.arrayBuffer()));
      } catch (error) {
        reject(error);
      }
    }, mimeType, format === "jpeg" ? 0.9 : undefined);
  });
}

async function rasterize(input) {
  const request = validateRequest(input);
  const pdfjsPath = pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.mjs")).href;
  const pdfjs = await import(pdfjsPath);
  pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.worker.mjs")).href;
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(await fs.readFile(request.pdfPath)),
    disableWorker: true,
    isEvalSupported: false,
    useSystemFonts: true,
  });
  const pdfDocument = await loadingTask.promise;
  try {
    if (pdfDocument.numPages < 1) throw new Error("The PowerPoint PDF render does not contain any pages.");
    if (pdfDocument.numPages > MAX_NATIVE_SLIDES) throw new Error("The presentation exceeds the 1,000-slide native render limit.");
    for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
      const page = await pdfDocument.getPage(pageNumber);
      try {
        const baseViewport = page.getViewport({ scale: 1 });
        const viewport = page.getViewport({ scale: request.width / baseViewport.width });
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(viewport.width);
        canvas.height = Math.round(viewport.height);
        const context = canvas.getContext("2d", { alpha: false });
        if (!context) throw new Error("Chromium could not create a slide canvas.");
        await page.render({ canvas, canvasContext: context, viewport, intent: "print", background: "rgb(255,255,255)" }).promise;
        const extension = request.format === "png" ? "png" : "jpg";
        await fs.writeFile(path.join(request.outputDirectory, `slide-${pageNumber}.${extension}`), await canvasBytes(canvas, request.format), { mode: 0o600 });
        canvas.width = 1;
        canvas.height = 1;
      } finally {
        page.cleanup();
      }
    }
    return { slideCount: pdfDocument.numPages };
  } finally {
    await pdfDocument.destroy();
  }
}

contextBridge.exposeInMainWorld("presentationStudioRasterize", rasterize);
