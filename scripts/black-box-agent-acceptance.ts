import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PRESENTATION_DESIGN_STANDARD } from "../src/lib/design-standard";

export interface BlackBoxAgentRunCase {
  id: string;
  sourceSlide: number;
  communicationJob: string;
  reviewFocus: string[];
}

export interface BlackBoxAgentRun {
  schema: "presentation-studio/black-box-agent-run";
  version: 1;
  id: string;
  preparedAt: string;
  designStandardVersion: string;
  sourceSha256: string;
  cases: BlackBoxAgentRunCase[];
  resultPath: string;
  rules: string[];
}

export interface BlackBoxAgentResult {
  schema: "presentation-studio/black-box-agent-result";
  version: 1;
  runId: string;
  completedAt: string;
  designStandardVersion: string;
  taskId?: string;
  sourceSha256: string;
  noSaveOrExport: boolean;
  cases: Array<{
    id: string;
    sourceSlide: number;
    status: "pass" | "hold";
    attempts: number;
    summary: string;
    defects: string[];
  }>;
}

function argument(name: string) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function object(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Expected a JSON object.");
  return value as Record<string, unknown>;
}

function absolutePath(value: string | undefined, label: string): string {
  if (!value || !path.isAbsolute(value)) throw new Error(`${label} must be an absolute path.`);
  return value;
}

function boundedText(value: unknown, label: string, maximum = 2_000): string {
  const text = String(value ?? "").trim();
  if (!text || text.length > maximum) throw new Error(`${label} must contain 1-${maximum} characters.`);
  return text;
}

function sha(value: unknown, label: string): string {
  const text = String(value ?? "");
  if (!/^[0-9a-f]{64}$/.test(text)) throw new Error(`${label} must be a lowercase SHA-256 digest.`);
  return text;
}

export function parseBlackBoxAgentResult(value: unknown, run: BlackBoxAgentRun): BlackBoxAgentResult {
  const raw = object(value);
  if (raw.schema !== "presentation-studio/black-box-agent-result" || raw.version !== 1) throw new Error("Use presentation-studio/black-box-agent-result version 1.");
  if (raw.runId !== run.id) throw new Error("The agent result does not match this run ID.");
  if (raw.designStandardVersion !== run.designStandardVersion) throw new Error("The agent tested a different Presentation Design Standard version.");
  if (sha(raw.sourceSha256, "sourceSha256") !== run.sourceSha256) throw new Error("The agent result does not match the source deck hash.");
  if (raw.noSaveOrExport !== true) throw new Error("The acceptance run must not save or export a user deliverable.");
  const results = Array.isArray(raw.cases) ? raw.cases : [];
  if (results.length !== run.cases.length) throw new Error("The agent result must contain exactly one result for every requested case.");
  const expected = new Map(run.cases.map((item) => [item.id, item]));
  const seen = new Set<string>();
  const cases = results.map((value, index) => {
    const item = object(value);
    const id = boundedText(item.id, `cases[${index}].id`, 120);
    const requested = expected.get(id);
    const sourceSlide = Number(item.sourceSlide);
    if (!requested || seen.has(id) || sourceSlide !== requested.sourceSlide) throw new Error(`Case ${id} is missing, duplicated, or bound to the wrong source slide.`);
    if (!['pass', 'hold'].includes(String(item.status))) throw new Error(`Case ${id} must be pass or hold.`);
    const attempts = Number(item.attempts);
    if (!Number.isInteger(attempts) || attempts < 1 || attempts > 3) throw new Error(`Case ${id} must record 1-3 attempts.`);
    seen.add(id);
    return {
      id,
      sourceSlide,
      status: item.status as "pass" | "hold",
      attempts,
      summary: boundedText(item.summary, `cases[${index}].summary`),
      defects: (Array.isArray(item.defects) ? item.defects : []).map((defect, defectIndex) => boundedText(defect, `cases[${index}].defects[${defectIndex}]`)).slice(0, 20),
    };
  });
  return {
    schema: "presentation-studio/black-box-agent-result",
    version: 1,
    runId: run.id,
    completedAt: boundedText(raw.completedAt, "completedAt", 80),
    designStandardVersion: run.designStandardVersion,
    taskId: raw.taskId ? boundedText(raw.taskId, "taskId", 120) : undefined,
    sourceSha256: run.sourceSha256,
    noSaveOrExport: true,
    cases,
  };
}

export async function prepareBlackBoxAgentRun(qualificationPath: string, outputRoot: string) {
  const qualification = object(JSON.parse(await fs.readFile(qualificationPath, "utf8")));
  if (qualification.schema !== "presentation-studio/private-golden-qualification" || qualification.version !== 2) throw new Error("Prepare black-box acceptance from a private-golden qualification version 2 ledger.");
  const source = object(qualification.source);
  const manifest = object(qualification.manifest);
  const rawCases = Array.isArray(qualification.cases) ? qualification.cases : [];
  if (!rawCases.length) throw new Error("The qualification ledger has no representative cases.");
  await fs.mkdir(outputRoot, { recursive: true });
  const id = `${boundedText(manifest.id, "manifest.id", 120)}-${PRESENTATION_DESIGN_STANDARD.version}`.replace(/[^a-z0-9._-]+/gi, "-").slice(0, 180);
  const resultPath = path.join(outputRoot, "fresh-agent-result.json");
  const run: BlackBoxAgentRun = {
    schema: "presentation-studio/black-box-agent-run",
    version: 1,
    id,
    preparedAt: new Date().toISOString(),
    designStandardVersion: PRESENTATION_DESIGN_STANDARD.version,
    sourceSha256: sha(source.sha256, "source.sha256"),
    cases: rawCases.map((value, index) => {
      const item = object(value);
      return {
        id: boundedText(item.id, `cases[${index}].id`, 120),
        sourceSlide: Number(item.sourceSlide),
        communicationJob: boundedText(item.communicationJob, `cases[${index}].communicationJob`, 500),
        reviewFocus: (Array.isArray(item.reviewFocus) ? item.reviewFocus : []).map(String),
      };
    }),
    resultPath,
    rules: [
      "Use only the installed Presentation Studio app, its MCP contract, the currently open source deck, and PowerPoint-native evidence.",
      "Do not inspect repository code, private-golden outputs, the human benchmark, or any prior task.",
      "Preserve exact content and source meaning. ORNL is the default. The approved title composition is sacred.",
      "Use at most three materially distinct native-reviewed attempts per case; record hold instead of looping or forcing a weaker candidate.",
      "Do not save or export a user deliverable.",
    ],
  };
  const runPath = path.join(outputRoot, "black-box-agent-run.json");
  await fs.writeFile(runPath, `${JSON.stringify(run, null, 2)}\n`);
  const cases = run.cases.map((item) => `${item.sourceSlide} (${item.id}: ${item.communicationJob})`).join(", ");
  const promptPath = path.join(outputRoot, "fresh-agent-prompt.txt");
  await fs.writeFile(promptPath, `Connect to the Presentation Studio MCP and work only in the currently open project. You are a context-free product acceptance agent: do not inspect repository code, other tasks, private-golden artifacts, or any human-cleaned benchmark.\n\nRead get_design_contract, get_app_status, list_decks, get_agent_runbook, and the deck work order. Test only these representative source slides: ${cases}. For each slide, follow the MCP composition plan, inspect immutable source pixels, build the exact editable PowerPoint candidate, inspect the full-size native result, run the Found issues -> Fixing -> Rechecking original intent loop, and apply the source-wins rule. Use no more than three materially distinct attempts. PASS only when the candidate is objectively clean and visually at least as strong as the source under the requested intervention; otherwise HOLD and report the product defect precisely. Do not save or export.\n\nAt the end write ${resultPath} as presentation-studio/black-box-agent-result version 1 for runId ${run.id}, designStandardVersion ${run.designStandardVersion}, sourceSha256 ${run.sourceSha256}, noSaveOrExport true, and one case result for every requested ID. Each case must include id, sourceSlide, status pass|hold, attempts 1-3, summary, and defects[]. Include your task ID when available.\n`);
  return { run, runPath, promptPath, resultPath };
}

export async function evaluateBlackBoxAgentRun(runPath: string, resultPath: string, outputRoot: string, baselinePath?: string) {
  const run = object(JSON.parse(await fs.readFile(runPath, "utf8"))) as unknown as BlackBoxAgentRun;
  const result = parseBlackBoxAgentResult(JSON.parse(await fs.readFile(resultPath, "utf8")), run);
  let baseline: BlackBoxAgentResult | undefined;
  if (baselinePath) {
    const rawBaseline = object(JSON.parse(await fs.readFile(baselinePath, "utf8")));
    baseline = parseBlackBoxAgentResult(rawBaseline, {
      ...run,
      id: boundedText(rawBaseline.runId, "baseline.runId", 180),
      designStandardVersion: boundedText(rawBaseline.designStandardVersion, "baseline.designStandardVersion", 180),
    });
  }
  const passCount = result.cases.filter((item) => item.status === "pass").length;
  const baselinePassCount = baseline?.cases.filter((item) => item.status === "pass").length;
  const trend = baselinePassCount === undefined ? "baseline" : passCount > baselinePassCount ? "progress" : passCount < baselinePassCount ? "regression" : "no-change";
  const summary = {
    schema: "presentation-studio/black-box-agent-acceptance",
    version: 1,
    evaluatedAt: new Date().toISOString(),
    run: { id: run.id, designStandardVersion: run.designStandardVersion, sourceSha256: run.sourceSha256 },
    result: { taskId: result.taskId, passCount, holdCount: result.cases.length - passCount, total: result.cases.length },
    baselinePassCount,
    trend,
    cases: result.cases,
    defects: [...new Set(result.cases.flatMap((item) => item.defects))],
    acceptanceRule: "Progress requires more context-free PASS cases without save/export, content loss, source-intent loss, or a weaker visual result.",
  };
  await fs.mkdir(outputRoot, { recursive: true });
  const summaryPath = path.join(outputRoot, "black-box-agent-acceptance.json");
  await fs.writeFile(summaryPath, `${JSON.stringify(summary, null, 2)}\n`);
  return { summary, summaryPath };
}

if (path.resolve(fileURLToPath(import.meta.url)) === path.resolve(process.argv[1] ?? "")) {
  const outputRoot = absolutePath(argument("output"), "--output");
  const qualificationPath = argument("qualification");
  const runPath = argument("run");
  const resultPath = argument("result");
  const task = qualificationPath
    ? prepareBlackBoxAgentRun(absolutePath(qualificationPath, "--qualification"), outputRoot)
    : runPath && resultPath
      ? evaluateBlackBoxAgentRun(absolutePath(runPath, "--run"), absolutePath(resultPath, "--result"), outputRoot, argument("baseline") ? absolutePath(argument("baseline"), "--baseline") : undefined)
      : Promise.reject(new Error("Prepare with --qualification and --output, or evaluate with --run, --result, and --output."));
  task.then((value) => process.stdout.write(`${JSON.stringify(value, null, 2)}\n`)).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
