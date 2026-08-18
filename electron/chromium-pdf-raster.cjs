const fsSync = require("node:fs");
const path = require("node:path");

const WORKER_HTML_PATH = path.join(__dirname, "pdf-raster-window.html");
const WORKER_PRELOAD_PATH = path.join(__dirname, "pdf-raster-preload.cjs");

function validateRasterRequest(input) {
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

function createChromiumPdfRasterizer(BrowserWindow, timeout = 180_000) {
  if (typeof BrowserWindow !== "function") throw new Error("A BrowserWindow constructor is required for Chromium PDF rendering.");
  return async function rasterizePdfWithChromium(input) {
    const request = validateRasterRequest(input);
    const workerWindow = new BrowserWindow({
      show: false,
      width: 320,
      height: 200,
      webPreferences: {
        preload: WORKER_PRELOAD_PATH,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: false,
        backgroundThrottling: false,
      },
    });
    try {
      await workerWindow.loadFile(WORKER_HTML_PATH);
      const renderPromise = workerWindow.webContents.executeJavaScript(`window.presentationStudioRasterize(${JSON.stringify(request)})`, true);
      const crashPromise = new Promise((_, reject) => {
        workerWindow.webContents.once("render-process-gone", (_event, details) => reject(new Error(`The isolated Chromium PDF renderer stopped unexpectedly (${details.reason}). Presentation Studio and the open project remain safe.`)));
      });
      const timeoutPromise = new Promise((_, reject) => {
        const timer = setTimeout(() => reject(new Error("The isolated Chromium PDF renderer timed out. Presentation Studio and the open project remain safe.")), timeout);
        timer.unref?.();
      });
      const result = await Promise.race([renderPromise, crashPromise, timeoutPromise]);
      if (!result || !Number.isInteger(result.slideCount) || result.slideCount < 1) throw new Error("The isolated Chromium PDF renderer did not report any slide images.");
    } finally {
      if (!workerWindow.isDestroyed()) workerWindow.destroy();
    }
  };
}

module.exports = {
  createChromiumPdfRasterizer,
  validateRasterRequest,
  WORKER_HTML_PATH,
  WORKER_PRELOAD_PATH,
};
