import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const RUNTIME_FILE_NAME = "mcp-runtime.json";
const REQUEST_TIMEOUT_MS = 18_000;
const NATIVE_RENDER_TIMEOUT_MS = 200_000;
const CENTRAL_BUILD_TIMEOUT_MS = 610_000;

export class PresentationAppUnavailableError extends Error {
  constructor(message = "Open Presentation Studio, then try this tool again.") {
    super(message);
    this.name = "PresentationAppUnavailableError";
  }
}

export function defaultRuntimeFilePath() {
  if (process.platform === "darwin") return path.join(os.homedir(), "Library", "Application Support", "Presentation Studio", RUNTIME_FILE_NAME);
  if (process.platform === "win32" && process.env.APPDATA) return path.join(process.env.APPDATA, "Presentation Studio", RUNTIME_FILE_NAME);
  return path.join(os.homedir(), ".presentation-studio", RUNTIME_FILE_NAME);
}

export function runtimeFilePath() {
  return process.env.PRESENTATION_STUDIO_MCP_RUNTIME_FILE || defaultRuntimeFilePath();
}

export function readRuntimeDescriptor(runtime = runtimeFilePath()) {
  let stats;
  try { stats = fs.lstatSync(runtime); } catch { throw new PresentationAppUnavailableError(); }
  if (!stats.isFile() || stats.isSymbolicLink()) throw new PresentationAppUnavailableError("The Presentation Studio connection file is not a regular local file.");
  if (process.platform !== "win32" && (stats.mode & 0o077) !== 0) throw new PresentationAppUnavailableError("The Presentation Studio connection file has unsafe permissions. Restart the app to repair it.");
  let descriptor;
  try { descriptor = JSON.parse(fs.readFileSync(runtime, "utf8")); } catch { throw new PresentationAppUnavailableError("The Presentation Studio connection file could not be read. Restart the app."); }
  if (descriptor?.version !== 1 || !Number.isInteger(descriptor.pid) || typeof descriptor.baseUrl !== "string" || typeof descriptor.token !== "string" || descriptor.token.length < 32) {
    throw new PresentationAppUnavailableError("The Presentation Studio connection file is invalid. Restart the app.");
  }
  let baseUrl;
  try { baseUrl = new URL(descriptor.baseUrl); } catch { throw new PresentationAppUnavailableError(); }
  if (baseUrl.protocol !== "http:" || baseUrl.hostname !== "127.0.0.1") throw new PresentationAppUnavailableError("Presentation Studio refused a desktop connection that was not loopback-only.");
  return { ...descriptor, baseUrl };
}

export class PresentationAppClient {
  constructor(options = {}) { this.runtimePath = options.runtimePath; }

  async command(operation, input = {}) {
    const descriptor = readRuntimeDescriptor(this.runtimePath);
    const url = new URL("command", descriptor.baseUrl);
    if (url.origin !== descriptor.baseUrl.origin) throw new Error("Presentation Studio refused a non-local request target.");
    let response;
    try {
      response = await fetch(url, {
        method: "POST",
        headers: { accept: "application/json", "content-type": "application/json", "x-presentation-studio-token": descriptor.token },
        body: JSON.stringify({ operation, input }),
        signal: AbortSignal.timeout(["build_studio_presentation", "run_deck_qualification"].includes(operation) ? CENTRAL_BUILD_TIMEOUT_MS : ["get_slide_render", "get_slide_render_comparison", "get_template_layout_render", "get_slide_design_work_order", "get_deck_design_work_order", "get_deck_contact_sheet", "get_qualification_contact_sheet", "get_slide_inspection_packet", "get_slide_measurements", "preview_studio_fresh_composition", "get_studio_slide_critique", "repair_studio_objective_issues", "record_studio_visual_critique", "record_proposal_visual_critique", "solve_and_stage_alignment", "solve_and_stage_distribution", "solve_and_stage_safe_region", "solve_and_stage_group_layout", "fit_scene_to_layout", "solve_and_stage_table_layout", "solve_and_stage_text_fit"].includes(operation) ? NATIVE_RENDER_TIMEOUT_MS : REQUEST_TIMEOUT_MS),
      });
    } catch { throw new PresentationAppUnavailableError(); }
    const body = response.headers.get("content-type")?.includes("application/json") ? await response.json() : { error: await response.text() };
    if (!response.ok) throw new Error(body && typeof body.error === "string" ? body.error.slice(0, 700) : `Presentation Studio returned status ${response.status}.`);
    return body.result;
  }
}
