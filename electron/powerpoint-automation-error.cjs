const { execFile } = require("node:child_process");
const { promisify } = require("node:util");

const execFileAsync = promisify(execFile);

// A timed-out or startup-window-blocked AppleScript can open a bridge copy
// before the caller binds targetPresentation. Close only the exact bridge
// path during Node cleanup so an orphan cannot block the next user open.
const POWERPOINT_CLOSE_EXACT_SCRIPT = `on run argv
  set sourcePath to item 1 of argv
  if application "Microsoft PowerPoint" is running then
    tell application "Microsoft PowerPoint"
      repeat with presentationIndex from 1 to (count of presentations)
        set candidatePresentation to presentation presentationIndex
        try
          if (full name of candidatePresentation as text) is sourcePath then
            close candidatePresentation saving no
            exit repeat
          end if
        end try
      end repeat
    end tell
  end if
  return "closed"
end run`;

async function closeExactPowerPointPresentation(sourcePath) {
  if (!sourcePath) return false;
  try {
    await execFileAsync("/usr/bin/osascript", ["-e", POWERPOINT_CLOSE_EXACT_SCRIPT, sourcePath], { timeout: 10_000, maxBuffer: 64 * 1024 });
    return true;
  } catch {
    return false;
  }
}

function automationErrorText(error) {
  return [error?.stderr, error?.stdout, error?.message, error?.code, error?.signal].filter(Boolean).join("\n");
}

function boundedPowerPointDetail(error) {
  const preferred = String(error?.stderr || error?.stdout || "").trim();
  if (preferred) return preferred.replace(/\s+/g, " ").slice(0, 420);
  const lines = String(error?.message || error || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !/^command failed:\s*\/usr\/bin\/osascript/i.test(line) && !/^(tell application|set |try$|end try$|end tell$|on error|open |close |save )/i.test(line));
  return (lines.at(-1) || "PowerPoint returned no additional diagnostic detail.").replace(/\s+/g, " ").slice(0, 420);
}

function describePowerPointAutomationError(error, action = "process") {
  const text = automationErrorText(error);
  if (/-1743|not authorized|not permitted to send apple events|automation permission/i.test(text) || /-128|user canceled|cancelled/i.test(text)) {
    return { status: "permission-required", reason: "automation-permission-required", message: `Allow Presentation Studio to control Microsoft PowerPoint, then retry native ${action}.` };
  }
  if (/mac is locked|session is locked|loginwindow|screen locked|automatic unlock could not unlock/i.test(text)) {
    return { status: "failed", reason: "mac-session-locked", message: `Unlock the Mac, leave Microsoft PowerPoint available, and retry native ${action}.` };
  }
  if (error?.killed || error?.code === "ETIMEDOUT" || error?.signal === "SIGTERM" || /-1712|apple event timed out|timed?\s*out/i.test(text)) {
    return { status: "failed", reason: "powerpoint-automation-timeout", message: `PowerPoint automation timed out during native ${action}. Unlock macOS, close any PowerPoint dialog, and retry.` };
  }
  if (/-9074|an error of type\s+-9074/i.test(text)) {
    return { status: "failed", reason: "powerpoint-startup-window-blocked", message: `Microsoft PowerPoint's Open new and recent files window blocked native ${action}. Close that window, then retry.` };
  }
  if (/modal|dialog|busy|another action|currently editing|cannot complete/i.test(text)) {
    return { status: "failed", reason: "powerpoint-busy", message: `Microsoft PowerPoint is waiting on another action or dialog. Finish or dismiss it, then retry native ${action}.` };
  }
  return { status: "failed", reason: "powerpoint-automation-failed", message: `Microsoft PowerPoint could not complete native ${action}: ${boundedPowerPointDetail(error)}` };
}

function classifyPowerPointAutomationError(error) {
  return describePowerPointAutomationError(error).status;
}

async function runPowerPointAutomationWithStartupRecovery({ action = "process", run, presentationCount, quit }) {
  try {
    return await run();
  } catch (error) {
    const diagnostic = describePowerPointAutomationError(error, action);
    if (diagnostic.reason !== "powerpoint-startup-window-blocked") throw error;

    let openPresentationCount;
    try {
      openPresentationCount = Number(await presentationCount());
    } catch {
      throw error;
    }
    if (!Number.isInteger(openPresentationCount) || openPresentationCount !== 0) throw error;

    try {
      await quit();
    } catch {
      throw error;
    }
    return run();
  }
}

module.exports = {
  POWERPOINT_CLOSE_EXACT_SCRIPT,
  automationErrorText,
  boundedPowerPointDetail,
  classifyPowerPointAutomationError,
  closeExactPowerPointPresentation,
  describePowerPointAutomationError,
  macSessionLocked,
  runPowerPointAutomationWithStartupRecovery,
};
const { execFileSync } = require("node:child_process");

function macSessionLocked(ioregOutput) {
  try {
    const output = ioregOutput === undefined
      ? execFileSync("/usr/sbin/ioreg", ["-n", "Root", "-d1", "-k", "IOConsoleUsers"], { encoding: "utf8", timeout: 2_000, maxBuffer: 1024 * 1024 })
      : String(ioregOutput);
    return /["']?CGSSessionScreenIsLocked["']?\s*=\s*Yes/i.test(output);
  } catch { return false; }
}
