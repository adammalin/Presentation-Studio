const fsSync = require("node:fs");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { execFile, spawnSync } = require("node:child_process");
const { createHash, randomUUID } = require("node:crypto");
const { promisify } = require("node:util");
const {
  classifyPowerPointAutomationError,
  closeExactPowerPointPresentation,
  describePowerPointAutomationError,
  macSessionLocked,
  runPowerPointAutomationWithStartupRecovery,
} = require("./powerpoint-automation-error.cjs");

const execFileAsync = promisify(execFile);
const POWERPOINT_MAC_PATH = "/Applications/Microsoft PowerPoint.app";
const MAX_NATIVE_SOURCE_BYTES = 1_250_000_000;
const MAX_NATIVE_RENDER_BYTES = 300_000_000;
const MAX_NATIVE_SLIDES = 1_000;

async function openPresentationCount() {
  const { stdout } = await execFileAsync("/usr/bin/osascript", ["-e", 'tell application "Microsoft PowerPoint" to return count of presentations'], { timeout: 10_000, maxBuffer: 64 * 1024 });
  return Number(String(stdout || "").trim());
}

async function quitPowerPointWithoutPresentations() {
  await execFileAsync("/usr/bin/osascript", ["-e", 'tell application "Microsoft PowerPoint" to quit'], { timeout: 15_000, maxBuffer: 64 * 1024 });
}

const POWERPOINT_RENDER_SCRIPT = `on run argv
  set sourcePath to item 1 of argv
  set outputPath to item 2 of argv
  set sourceFile to POSIX file sourcePath
  set outputFile to POSIX file outputPath
  tell application "Microsoft PowerPoint"
    set targetPresentation to missing value
    try
      open sourceFile
      repeat with bindAttempt from 1 to 40
        repeat with presentationIndex from 1 to (count of presentations)
          set candidatePresentation to presentation presentationIndex
          try
            if (full name of candidatePresentation as text) is sourcePath then
              set targetPresentation to candidatePresentation
              exit repeat
            end if
          end try
        end repeat
        if targetPresentation is not missing value then exit repeat
        delay 0.1
      end repeat
      if targetPresentation is missing value then error "PowerPoint did not open the requested render copy."
      -- PowerPoint can expose the presentation object before its text and
      -- linked image layout is ready for PDF output. Without this brief
      -- stabilization window, a newly generated slide can intermittently
      -- export with its first text shape missing even though the PPTX package
      -- and a subsequent export are correct.
      delay 1
      save targetPresentation in outputFile as save as PDF
      close targetPresentation saving no
    on error renderError number renderErrorNumber
      if targetPresentation is not missing value then
        try
          close targetPresentation saving no
        end try
      else
        -- The open command may have created the presentation before a
        -- startup-window error prevented the exact binding above.
        repeat with presentationIndex from 1 to (count of presentations)
          set candidatePresentation to presentation presentationIndex
          try
            if (full name of candidatePresentation as text) is sourcePath then
              close candidatePresentation saving no
              exit repeat
            end if
          end try
        end repeat
      end if
      error renderError number renderErrorNumber
    end try
  end tell
  return "rendered|" & version of application "Microsoft PowerPoint"
end run`;

function executableOnPath(name, environmentPath = process.env.PATH || "") {
  for (const directory of environmentPath.split(path.delimiter).filter(Boolean)) {
    const candidate = path.join(directory, name);
    try {
      fsSync.accessSync(candidate, fsSync.constants.X_OK);
      return candidate;
    } catch { /* Keep looking. */ }
  }
  return null;
}

function validatePdfRasterizer(candidate) {
  if (!candidate) return null;
  try {
    fsSync.accessSync(candidate, fsSync.constants.X_OK);
    // Invoke the canonical target rather than a node_modules/.bin symlink. The
    // bundled Poppler launchers locate their libraries relative to BASH_SOURCE;
    // executing a symlink makes that relative lookup start in the consumer
    // project's .bin directory and breaks otherwise valid installations.
    const canonicalPath = fsSync.realpathSync(candidate);
    fsSync.accessSync(canonicalPath, fsSync.constants.X_OK);
    const probe = spawnSync(canonicalPath, ["-v"], {
      encoding: "utf8",
      stdio: "ignore",
      timeout: 5_000,
    });
    return probe.status === 0 && !probe.error ? canonicalPath : null;
  } catch {
    return null;
  }
}

function resolvePdfRasterizer() {
  const configured = process.env.PRESENTATION_STUDIO_PDFTOPPM;
  const candidates = [
    configured,
    executableOnPath("pdftoppm"),
    "/opt/homebrew/bin/pdftoppm",
    "/usr/local/bin/pdftoppm",
  ].filter(Boolean);
  for (const candidate of candidates) {
    const validated = validatePdfRasterizer(candidate);
    if (validated) return validated;
  }
  return null;
}

function nativeRenderCapabilities(platform = process.platform) {
  const rasterizerPath = resolvePdfRasterizer();
  if (platform !== "darwin") {
    return {
      available: false,
      renderer: "studio-approximate",
      reason: platform === "win32" ? "The Windows PowerPoint render bridge is not implemented yet." : "PowerPoint-native rendering currently requires macOS.",
      powerPointInstalled: false,
      rasterizerAvailable: Boolean(rasterizerPath),
      sessionLocked: false,
    };
  }
  const powerPointInstalled = fsSync.existsSync(POWERPOINT_MAC_PATH);
  const sessionLocked = powerPointInstalled && macSessionLocked();
  return {
    available: powerPointInstalled && Boolean(rasterizerPath) && !sessionLocked,
    renderer: powerPointInstalled && rasterizerPath ? "powerpoint-native" : "studio-approximate",
    reason: !powerPointInstalled ? "Microsoft PowerPoint is not installed." : !rasterizerPath ? "The local PDF rasterizer is unavailable." : sessionLocked ? "Unlock the Mac to enable PowerPoint-native rendering." : undefined,
    powerPointInstalled,
    rasterizerAvailable: Boolean(rasterizerPath),
    sessionLocked,
  };
}

function jpegDimensions(bytes) {
  if (!Buffer.isBuffer(bytes)) bytes = Buffer.from(bytes);
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) throw new Error("The native renderer produced an invalid JPEG image.");
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) { offset += 1; continue; }
    const marker = bytes[offset + 1];
    offset += 2;
    if (marker === 0xd8 || marker === 0xd9 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 2 > bytes.length) break;
    const length = bytes.readUInt16BE(offset);
    if (length < 2 || offset + length > bytes.length) break;
    if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
      return { height: bytes.readUInt16BE(offset + 3), width: bytes.readUInt16BE(offset + 5) };
    }
    offset += length;
  }
  throw new Error("The native renderer could not read the JPEG dimensions.");
}

function pngDimensions(bytes) {
  if (!Buffer.isBuffer(bytes)) bytes = Buffer.from(bytes);
  if (bytes.length < 24 || bytes.subarray(0, 8).toString("hex") !== "89504e470d0a1a0a") throw new Error("The native renderer produced an invalid PNG image.");
  return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) };
}

function slideNumberFromFile(fileName) {
  const match = fileName.match(/-(\d+)\.(?:jpe?g|png)$/i);
  return match ? Number(match[1]) : Number.NaN;
}

async function renderPowerPointNative({ bytes: inputBytes, name = "presentation.pptx", homePath = os.homedir(), width = 1400, format = "jpeg" }) {
  const capabilities = nativeRenderCapabilities();
  if (capabilities.sessionLocked) return { status: "failed", renderer: "powerpoint-native", authoritative: false, reason: "mac-session-locked", slides: [], warnings: ["Unlock the Mac, leave Microsoft PowerPoint available, and retry native rendering."] };
  if (!capabilities.available) {
    return { status: "unavailable", ...capabilities, authoritative: false, slides: [], warnings: capabilities.reason ? [capabilities.reason] : [] };
  }
  const bytes = Buffer.from(inputBytes ?? []);
  if (bytes.length === 0) throw new Error("PowerPoint-native rendering received an empty presentation.");
  if (bytes.length > MAX_NATIVE_SOURCE_BYTES) throw new Error("The presentation exceeds the native render safety limit.");
  if (!/^PK/.test(bytes.subarray(0, 2).toString("ascii"))) throw new Error("PowerPoint-native rendering requires a valid PPTX package.");
  if (!Number.isInteger(width) || width < 800 || width > 3_000) throw new Error("Native render width must be an integer from 800 to 3,000 pixels.");
  if (!["jpeg", "png"].includes(format)) throw new Error("Native render format must be JPEG or PNG.");

  const digest = createHash("sha256").update(bytes).digest("hex");
  const jobId = randomUUID();
  const bridgeRoot = path.join(homePath, "Library", "Containers", "com.microsoft.Powerpoint", "Data", "Documents", "Presentation Studio Render Bridge");
  const jobRoot = path.join(bridgeRoot, jobId);
  const safeStem = path.basename(name, path.extname(name)).replace(/[^a-z0-9._-]+/gi, "-").slice(0, 70) || "presentation";
  const sourcePath = path.join(jobRoot, `${safeStem}-${jobId}.pptx`);
  const pdfPath = path.join(jobRoot, `${safeStem}-${jobId}.pdf`);
  const imagePrefix = path.join(jobRoot, "slide");
  try {
    await fs.mkdir(jobRoot, { recursive: true, mode: 0o700 });
    await fs.writeFile(sourcePath, bytes, { mode: 0o600 });
    try {
      var automation = await runPowerPointAutomationWithStartupRecovery({
        action: "rendering",
        run: () => execFileAsync("/usr/bin/osascript", ["-e", POWERPOINT_RENDER_SCRIPT, sourcePath, pdfPath], { timeout: 180_000, maxBuffer: 1024 * 1024 }),
        presentationCount: openPresentationCount,
        quit: quitPowerPointWithoutPresentations,
      });
    } catch (error) {
      const diagnostic = describePowerPointAutomationError(error, "rendering");
      return {
        status: diagnostic.status,
        renderer: "powerpoint-native",
        authoritative: false,
        sourceSha256: digest,
        slides: [],
        reason: diagnostic.reason,
        warnings: [diagnostic.message],
      };
    }
    const pdfStat = await fs.stat(pdfPath).catch(() => null);
    if (!pdfStat?.isFile() || pdfStat.size === 0) throw new Error("Microsoft PowerPoint did not create the expected native PDF render.");
    const rasterArguments = format === "png"
      ? ["-png", "-scale-to-x", String(width), "-scale-to-y", "-1", pdfPath, imagePrefix]
      : ["-jpeg", "-scale-to-x", String(width), "-scale-to-y", "-1", "-jpegopt", "quality=90,progressive=y,optimize=y", pdfPath, imagePrefix];
    await execFileAsync(capabilities.rasterizerPath || resolvePdfRasterizer(), rasterArguments, { timeout: 180_000, maxBuffer: 1024 * 1024 });
    const fileNames = (await fs.readdir(jobRoot)).filter((fileName) => format === "png" ? /^slide-\d+\.png$/i.test(fileName) : /^slide-\d+\.jpe?g$/i.test(fileName)).sort((left, right) => slideNumberFromFile(left) - slideNumberFromFile(right));
    if (fileNames.length === 0) throw new Error("The local rasterizer did not create any slide images.");
    if (fileNames.length > MAX_NATIVE_SLIDES) throw new Error("The presentation exceeds the 1,000-slide native render limit.");
    const slides = [];
    let totalBytes = 0;
    for (const fileName of fileNames) {
      const data = await fs.readFile(path.join(jobRoot, fileName));
      totalBytes += data.length;
      if (totalBytes > MAX_NATIVE_RENDER_BYTES) throw new Error("The native slide images exceed the 300 MB in-memory safety limit.");
      const dimensions = format === "png" ? pngDimensions(data) : jpegDimensions(data);
      slides.push({ number: slideNumberFromFile(fileName), mimeType: format === "png" ? "image/png" : "image/jpeg", width: dimensions.width, height: dimensions.height, sha256: createHash("sha256").update(data).digest("hex"), bytes: new Uint8Array(data) });
    }
    return {
      status: "ready",
      renderer: "powerpoint-native",
      pipeline: `powerpoint-save-as-pdf+local-pdf-raster:${format}:${width}px`,
      powerPointVersion: automation?.stdout?.trim().split("|")[1] || undefined,
      authoritative: true,
      sourceSha256: digest,
      generatedAt: new Date().toISOString(),
      slideCount: slides.length,
      slides,
      warnings: [],
    };
  } finally {
    await closeExactPowerPointPresentation(sourcePath);
    await fs.rm(jobRoot, { recursive: true, force: true }).catch(() => undefined);
  }
}

module.exports = {
  POWERPOINT_RENDER_SCRIPT,
  classifyPowerPointAutomationError,
  jpegDimensions,
  pngDimensions,
  nativeRenderCapabilities,
  renderPowerPointNative,
  resolvePdfRasterizer,
  slideNumberFromFile,
  validatePdfRasterizer,
};
