const fsSync = require("node:fs");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { execFile } = require("node:child_process");
const { createHash, randomUUID } = require("node:crypto");
const { promisify } = require("node:util");
const {
  classifyPowerPointAutomationError: classifyMeasurementError,
  closeExactPowerPointPresentation,
  describePowerPointAutomationError,
  macSessionLocked,
  runPowerPointAutomationWithStartupRecovery,
} = require("./powerpoint-automation-error.cjs");

const execFileAsync = promisify(execFile);
const POWERPOINT_MAC_PATH = "/Applications/Microsoft PowerPoint.app";
const MAX_NATIVE_SOURCE_BYTES = 1_250_000_000;
const MAX_MEASUREMENT_BYTES = 64 * 1024 * 1024;

async function openPresentationCount() {
  const { stdout } = await execFileAsync("/usr/bin/osascript", ["-e", 'tell application "Microsoft PowerPoint" to return count of presentations'], { timeout: 10_000, maxBuffer: 64 * 1024 });
  return Number(String(stdout || "").trim());
}

async function quitPowerPointWithoutPresentations() {
  await execFileAsync("/usr/bin/osascript", ["-e", 'tell application "Microsoft PowerPoint" to quit'], { timeout: 15_000, maxBuffer: 64 * 1024 });
}

// PowerPoint is authoritative for rendered text geometry. The bridge deliberately
// records measurements only: it never returns visible copy or file bytes.
const POWERPOINT_MEASUREMENT_BODY = `
    set measurementLines to {}
    set end of measurementLines to "DECK\t" & (version of application "Microsoft PowerPoint") & "\t" & (count of slides of targetPresentation)
    repeat with slideIndex from 1 to (count of slides of targetPresentation)
      set measurementStage to "slide " & slideIndex & " access"
      set targetSlide to slide slideIndex of targetPresentation
      set measurementStage to "slide " & slideIndex & " shape inventory"
      set shapeCount to count of shapes of targetSlide
      set end of measurementLines to "SLIDE\t" & slideIndex & "\t" & shapeCount
      repeat with shapeIndex from 1 to shapeCount
        set measurementStage to "slide " & slideIndex & " shape " & shapeIndex & " measurement"
        try
          set targetShape to shape shapeIndex of targetSlide
          set shapeHasText to has text frame of targetShape
          set shapeHasTable to has table of targetShape
          set nativeShapeId to ""
          set nativeShapeName to ""
          try
            set nativeShapeId to id of targetShape as text
          end try
          try
            set nativeShapeName to name of targetShape as text
          end try
          set end of measurementLines to "SHAPE\t" & slideIndex & "\t" & shapeIndex & "\t" & (z order position of targetShape) & "\t" & (left position of targetShape) & "\t" & (top of targetShape) & "\t" & (width of targetShape) & "\t" & (height of targetShape) & "\t" & (rotation of targetShape) & "\t" & shapeHasText & "\t" & shapeHasTable & "\t" & nativeShapeId & "\t" & nativeShapeName
          if shapeHasText then
            try
              set targetFrame to text frame of targetShape
              set targetRange to text range of targetFrame
              set lineCount to count of lines of targetRange
              set end of measurementLines to "TEXT\t" & slideIndex & "\t" & shapeIndex & "\t" & (margin left of targetFrame) & "\t" & (margin right of targetFrame) & "\t" & (margin top of targetFrame) & "\t" & (margin bottom of targetFrame) & "\t" & (left bounds of targetRange) & "\t" & (top bounds of targetRange) & "\t" & (bounds width of targetRange) & "\t" & (bounds height of targetRange) & "\t" & (text length of targetRange) & "\t" & lineCount & "\t" & (vertical anchor of targetFrame)
            end try
          end if
          if shapeHasTable then
            try
              set targetTable to table object of targetShape
              set rowCount to count of rows of targetTable
              set columnCount to count of columns of targetTable
              set end of measurementLines to "TABLE\t" & slideIndex & "\t" & shapeIndex & "\t" & rowCount & "\t" & columnCount
              repeat with rowIndex from 1 to rowCount
                set end of measurementLines to "ROW\t" & slideIndex & "\t" & shapeIndex & "\t" & rowIndex & "\t" & (height of row rowIndex of targetTable)
              end repeat
              repeat with columnIndex from 1 to columnCount
                set end of measurementLines to "COL\t" & slideIndex & "\t" & shapeIndex & "\t" & columnIndex & "\t" & (width of column columnIndex of targetTable)
              end repeat
              repeat with rowIndex from 1 to rowCount
                repeat with columnIndex from 1 to columnCount
                  try
                    set targetCell to get cell from targetTable row rowIndex column columnIndex
                    set cellShape to shape of targetCell
                    set cellFrame to text frame of cellShape
                    set cellRange to text range of cellFrame
                    set cellLineCount to count of lines of cellRange
                    set end of measurementLines to "CELL\t" & slideIndex & "\t" & shapeIndex & "\t" & rowIndex & "\t" & columnIndex & "\t" & (left position of cellShape) & "\t" & (top of cellShape) & "\t" & (width of cellShape) & "\t" & (height of cellShape) & "\t" & (margin left of cellFrame) & "\t" & (margin right of cellFrame) & "\t" & (margin top of cellFrame) & "\t" & (margin bottom of cellFrame) & "\t" & (left bounds of cellRange) & "\t" & (top bounds of cellRange) & "\t" & (bounds width of cellRange) & "\t" & (bounds height of cellRange) & "\t" & (text length of cellRange) & "\t" & cellLineCount & "\t" & (vertical anchor of cellFrame)
                  on error cellError
                    set end of measurementLines to "CELL_ERROR\t" & slideIndex & "\t" & shapeIndex & "\t" & rowIndex & "\t" & columnIndex
                  end try
                end repeat
              end repeat
            end try
          end if
        on error shapeError
          set end of measurementLines to "SHAPE_ERROR\t" & slideIndex & "\t" & shapeIndex
        end try
      end repeat
    end repeat
    set measurementStage to "serialize measurement rows"
    set oldDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to linefeed
    set measurementText to measurementLines as text
    set AppleScript's text item delimiters to oldDelimiters`;

const POWERPOINT_MEASUREMENT_SCRIPT = `on run argv
  set sourcePath to item 1 of argv
  set measurementPath to item 2 of argv
  set sourceFile to POSIX file sourcePath
  set measurementFile to POSIX file measurementPath
  tell application "Microsoft PowerPoint"
    set targetPresentation to missing value
    set measurementStage to "open source copy"
    try
      open sourceFile
      set measurementStage to "bind exact presentation"
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
      if targetPresentation is missing value then error "PowerPoint did not open the requested measurement copy."
      -- Keep measurement synchronized with the authoritative render path.
      -- PowerPoint may bind the presentation before it has completed the
      -- initial text/image layout pass for a freshly generated PPTX.
      delay 1
${POWERPOINT_MEASUREMENT_BODY}
      set measurementStage to "close measured presentation"
      close targetPresentation saving no
    on error measurementError number measurementErrorNumber
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
      error "PowerPoint-native measurement failed during " & measurementStage & ": " & measurementError number measurementErrorNumber
    end try
  end tell
  set measurementHandle to open for access measurementFile with write permission
  try
    set eof measurementHandle to 0
    write measurementText to measurementHandle as «class utf8»
    close access measurementHandle
  on error writeError
    try
      close access measurementHandle
    end try
    error writeError
  end try
  return "measured"
end run`;

function parseBoolean(value) {
  return /^true$/i.test(String(value));
}

function finiteNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function bounds(left, top, width, height) {
  const values = [left, top, width, height].map(finiteNumber);
  if (values.some((value) => value === undefined)) return undefined;
  return { left: values[0], top: values[1], width: values[2], height: values[3] };
}

function margins(left, right, top, bottom) {
  const values = [left, right, top, bottom].map(finiteNumber);
  if (values.some((value) => value === undefined)) return undefined;
  return { left: values[0], right: values[1], top: values[2], bottom: values[3] };
}

function parsePowerPointMeasurement(text, { sourceSha256, generatedAt = new Date().toISOString() } = {}) {
  const result = {
    status: "ready",
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    sourceSha256,
    generatedAt,
    powerPointVersion: undefined,
    slideCount: 0,
    slides: [],
    warnings: [],
  };
  const slideMap = new Map();
  const shapeMap = new Map();
  for (const rawLine of String(text || "").split(/\r?\n/)) {
    const line = rawLine.replace(/^\uFEFF/, "").trim();
    if (!line) continue;
    const fields = line.split("\t");
    const kind = fields[0];
    if (kind === "DECK") {
      result.powerPointVersion = fields[1] || undefined;
      result.slideCount = finiteNumber(fields[2]) || 0;
      continue;
    }
    const slideNumber = finiteNumber(fields[1]);
    if (!slideNumber) continue;
    let slide = slideMap.get(slideNumber);
    if (!slide) {
      slide = { number: slideNumber, shapeCount: 0, shapes: [] };
      slideMap.set(slideNumber, slide);
      result.slides.push(slide);
    }
    if (kind === "SLIDE") {
      slide.shapeCount = finiteNumber(fields[2]) || 0;
      continue;
    }
    const shapeIndex = finiteNumber(fields[2]);
    if (!shapeIndex) continue;
    const key = `${slideNumber}:${shapeIndex}`;
    if (kind === "SHAPE") {
      const shape = {
        slideNumber,
        shapeIndex,
        zOrder: finiteNumber(fields[3]) || shapeIndex,
        boundsPt: bounds(fields[4], fields[5], fields[6], fields[7]),
        rotation: finiteNumber(fields[8]) || 0,
        hasTextFrame: parseBoolean(fields[9]),
        hasTable: parseBoolean(fields[10]),
        nativeShapeId: fields[11] || undefined,
        name: fields.slice(12).join("\t") || undefined,
      };
      shapeMap.set(key, shape);
      slide.shapes.push(shape);
      continue;
    }
    const shape = shapeMap.get(key);
    if (!shape) continue;
    if (kind === "TEXT") {
      shape.text = {
        marginsPt: margins(fields[3], fields[4], fields[5], fields[6]),
        renderedBoundsPt: bounds(fields[7], fields[8], fields[9], fields[10]),
        coordinateSpace: "slide",
        textLength: finiteNumber(fields[11]) || 0,
        lineCount: finiteNumber(fields[12]) || 0,
        verticalAnchor: fields[13] || "unknown",
      };
      continue;
    }
    if (kind === "TABLE") {
      shape.table = {
        rowCount: finiteNumber(fields[3]) || 0,
        columnCount: finiteNumber(fields[4]) || 0,
        rowHeightsPt: [],
        columnWidthsPt: [],
        cells: [],
      };
      continue;
    }
    if (!shape.table) continue;
    if (kind === "ROW") {
      shape.table.rowHeightsPt[(finiteNumber(fields[3]) || 1) - 1] = finiteNumber(fields[4]) || 0;
    } else if (kind === "COL") {
      shape.table.columnWidthsPt[(finiteNumber(fields[3]) || 1) - 1] = finiteNumber(fields[4]) || 0;
    } else if (kind === "CELL") {
      const cellBounds = bounds(fields[5], fields[6], fields[7], fields[8]);
      const tableRelativeTextBounds = bounds(fields[13], fields[14], fields[15], fields[16]);
      // PowerPoint's cell TextRange bounds are relative to the table shape,
      // while the cell shape bounds are slide-relative. Normalize the text
      // bounds to the cell before exposing the declared cell-relative packet.
      // Treating the raw table-relative left/top as cell-relative produces
      // increasingly negative clearance values in later columns.
      const cellRelativeTextBounds = cellBounds && tableRelativeTextBounds && shape.boundsPt ? {
        left: shape.boundsPt.left + tableRelativeTextBounds.left - cellBounds.left,
        top: shape.boundsPt.top + tableRelativeTextBounds.top - cellBounds.top,
        width: tableRelativeTextBounds.width,
        height: tableRelativeTextBounds.height,
      } : tableRelativeTextBounds;
      shape.table.cells.push({
        row: finiteNumber(fields[3]) || 0,
        column: finiteNumber(fields[4]) || 0,
        boundsPt: cellBounds,
        marginsPt: margins(fields[9], fields[10], fields[11], fields[12]),
        renderedTextBoundsPt: cellRelativeTextBounds,
        textCoordinateSpace: "cell-relative",
        textLength: finiteNumber(fields[17]) || 0,
        lineCount: finiteNumber(fields[18]) || 0,
        verticalAnchor: fields[19] || "unknown",
      });
    } else if (kind === "CELL_ERROR") {
      result.warnings.push(`PowerPoint could not measure slide ${slideNumber}, table shape ${shapeIndex}, cell ${fields[3]}:${fields[4]}.`);
    }
  }
  result.slides.sort((left, right) => left.number - right.number);
  for (const slide of result.slides) slide.shapes.sort((left, right) => left.shapeIndex - right.shapeIndex);
  if (!result.slideCount) result.slideCount = result.slides.length;
  if (result.slides.length === 0) throw new Error("PowerPoint returned no native slide measurements.");
  return result;
}

function nativeMeasurementCapabilities(platform = process.platform) {
  if (platform !== "darwin") {
    return { available: false, adapter: platform === "win32" ? "windows-powerpoint-com-pending" : "ooxml-fallback", reason: platform === "win32" ? "The Windows PowerPoint measurement adapter is not implemented yet." : "Native PowerPoint measurement currently requires macOS.", sessionLocked: false };
  }
  const installed = fsSync.existsSync(POWERPOINT_MAC_PATH);
  const sessionLocked = installed && macSessionLocked();
  return installed
    ? { available: !sessionLocked, adapter: "macos-powerpoint-applescript", reason: sessionLocked ? "Unlock the Mac to enable PowerPoint-native measurement." : undefined, sessionLocked }
    : { available: false, adapter: "ooxml-fallback", reason: "Microsoft PowerPoint is not installed.", sessionLocked: false };
}

async function measurePowerPointNative({ bytes: inputBytes, name = "presentation.pptx", homePath = os.homedir() }) {
  const capabilities = nativeMeasurementCapabilities();
  if (capabilities.sessionLocked) return { status: "failed", reason: "mac-session-locked", adapter: capabilities.adapter, authority: "unknown", slides: [], warnings: ["Unlock the Mac, leave Microsoft PowerPoint available, and retry native measurement."] };
  if (!capabilities.available) return { status: "unavailable", authority: "unknown", slides: [], warnings: capabilities.reason ? [capabilities.reason] : [], ...capabilities };
  const bytes = Buffer.from(inputBytes ?? []);
  if (bytes.length === 0) throw new Error("PowerPoint-native measurement received an empty presentation.");
  if (bytes.length > MAX_NATIVE_SOURCE_BYTES) throw new Error("The presentation exceeds the native measurement safety limit.");
  if (!/^PK/.test(bytes.subarray(0, 2).toString("ascii"))) throw new Error("PowerPoint-native measurement requires a valid PPTX package.");
  const digest = createHash("sha256").update(bytes).digest("hex");
  const jobId = randomUUID();
  const bridgeRoot = path.join(homePath, "Library", "Containers", "com.microsoft.Powerpoint", "Data", "Documents", "Presentation Studio Measurement Bridge");
  const jobRoot = path.join(bridgeRoot, jobId);
  const safeStem = path.basename(name, path.extname(name)).replace(/[^a-z0-9._-]+/gi, "-").slice(0, 70) || "presentation";
  const sourcePath = path.join(jobRoot, `${safeStem}-${jobId}.pptx`);
  const measurementPath = path.join(jobRoot, "measurement.tsv");
  try {
    await fs.mkdir(jobRoot, { recursive: true, mode: 0o700 });
    await fs.writeFile(sourcePath, bytes, { mode: 0o600 });
    try {
      await runPowerPointAutomationWithStartupRecovery({
        action: "measurement",
        // A full 100-200 slide export-acceptance pass can legitimately exceed
        // three minutes. Interactive slide tools isolate one slide before
        // calling this adapter, while whole-deck qualification gets a bounded
        // ten-minute lane instead of being misreported as unavailable.
        run: () => execFileAsync("/usr/bin/osascript", ["-e", POWERPOINT_MEASUREMENT_SCRIPT, sourcePath, measurementPath], { timeout: 600_000, maxBuffer: 2 * 1024 * 1024 }),
        presentationCount: openPresentationCount,
        quit: quitPowerPointWithoutPresentations,
      });
    } catch (error) {
      const diagnostic = describePowerPointAutomationError(error, "measurement");
      return { status: diagnostic.status, reason: diagnostic.reason, adapter: capabilities.adapter, authority: "unknown", sourceSha256: digest, slides: [], warnings: [diagnostic.message] };
    }
    const stat = await fs.stat(measurementPath).catch(() => null);
    if (!stat?.isFile() || stat.size === 0) throw new Error("Microsoft PowerPoint did not create the expected native measurement file.");
    if (stat.size > MAX_MEASUREMENT_BYTES) throw new Error("The native measurement result exceeds the 64 MB safety limit.");
    return parsePowerPointMeasurement(await fs.readFile(measurementPath, "utf8"), { sourceSha256: digest });
  } finally {
    await closeExactPowerPointPresentation(sourcePath);
    await fs.rm(jobRoot, { recursive: true, force: true }).catch(() => undefined);
  }
}

module.exports = {
  POWERPOINT_MEASUREMENT_BODY,
  POWERPOINT_MEASUREMENT_SCRIPT,
  classifyMeasurementError,
  measurePowerPointNative,
  nativeMeasurementCapabilities,
  parsePowerPointMeasurement,
};
