import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { sha256 } from "../src/lib/hash";
import { qualifyStudioWebBenchmark } from "./qualify-studio-web-benchmark";

interface PrivateGoldenManifest {
  schema: "presentation-studio/private-golden-manifest";
  version: 1 | 2;
  id: string;
  source: { path: string; sha256: string };
  benchmark: { path: string; sha256: string };
  template?: { path: string; sha256: string };
  designMode?: "shared" | "template";
  cases: Array<{
    id: string;
    sourceSlide: number;
    benchmarkSlides: number[];
    communicationJob: string;
    reviewFocus: Array<"hierarchy" | "layout-balance" | "table-quality" | "template-fidelity" | "figure-clarity" | "editability" | "source-intent">;
  }>;
}

function argument(name: string) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function absoluteFile(value: unknown, label: string) {
  if (typeof value !== "string" || !path.isAbsolute(value)) throw new Error(`${label} must be an absolute local file path.`);
  return value;
}

function digest(value: unknown, label: string) {
  if (typeof value !== "string" || !/^[0-9a-f]{64}$/.test(value)) throw new Error(`${label} must be a lowercase SHA-256 digest.`);
  return value;
}

export function parsePrivateGoldenManifest(value: unknown): PrivateGoldenManifest {
  if (!value || typeof value !== "object") throw new Error("The private golden manifest must be a JSON object.");
  const raw = value as Record<string, unknown>;
  if (raw.schema !== "presentation-studio/private-golden-manifest" || ![1, 2].includes(Number(raw.version))) throw new Error("Use presentation-studio/private-golden-manifest version 1 or 2.");
  if (typeof raw.id !== "string" || !/^[a-z0-9._-]{1,120}$/i.test(raw.id)) throw new Error("The private golden manifest requires a stable ID.");
  const source = raw.source as Record<string, unknown>;
  const benchmark = raw.benchmark as Record<string, unknown>;
  const template = raw.template as Record<string, unknown> | undefined;
  const designMode = raw.designMode === undefined ? "shared" : raw.designMode;
  if (!['shared', 'template'].includes(String(designMode))) throw new Error("designMode must be shared or template.");
  if (designMode === "template" && !template) throw new Error("Template designMode requires an authorized template path and SHA-256.");
  const cases = Array.isArray(raw.cases) ? raw.cases : [];
  if (!cases.length || cases.length > 12) throw new Error("Choose 1–12 representative private golden cases.");
  const ids = new Set<string>();
  const slides = new Set<number>();
  const allowedFocus = new Set(["hierarchy", "layout-balance", "table-quality", "template-fidelity", "figure-clarity", "editability", "source-intent"]);
  const normalizedCases = cases.map((item, index) => {
    if (!item || typeof item !== "object") throw new Error(`Golden case ${index + 1} is invalid.`);
    const candidate = item as Record<string, unknown>;
    const id = String(candidate.id ?? "");
    const sourceSlide = Number(candidate.sourceSlide);
    const benchmarkSlides = Array.isArray(candidate.benchmarkSlides) ? candidate.benchmarkSlides.map(Number) : [];
    const communicationJob = String(candidate.communicationJob ?? "").trim();
    const reviewFocus = Array.isArray(candidate.reviewFocus) ? candidate.reviewFocus.map(String) : [];
    if (!/^[a-z0-9._-]{1,120}$/i.test(id) || ids.has(id)) throw new Error(`Golden case ${index + 1} needs a unique stable ID.`);
    if (!Number.isInteger(sourceSlide) || sourceSlide < 1 || slides.has(sourceSlide)) throw new Error(`Golden case ${id} needs a unique positive source slide.`);
    if (!benchmarkSlides.length || benchmarkSlides.some((slide) => !Number.isInteger(slide) || slide < 1)) throw new Error(`Golden case ${id} needs one or more positive benchmark slides.`);
    if (!communicationJob || communicationJob.length > 500) throw new Error(`Golden case ${id} needs a bounded communication job.`);
    if (!reviewFocus.length || reviewFocus.some((focus) => !allowedFocus.has(focus))) throw new Error(`Golden case ${id} has an invalid review focus.`);
    ids.add(id);
    slides.add(sourceSlide);
    return { id, sourceSlide, benchmarkSlides, communicationJob, reviewFocus: reviewFocus as PrivateGoldenManifest["cases"][number]["reviewFocus"] };
  });
  return {
    schema: "presentation-studio/private-golden-manifest",
    version: Number(raw.version) as 1 | 2,
    id: raw.id,
    source: { path: absoluteFile(source.path, "source.path"), sha256: digest(source.sha256, "source.sha256") },
    benchmark: { path: absoluteFile(benchmark.path, "benchmark.path"), sha256: digest(benchmark.sha256, "benchmark.sha256") },
    template: template ? { path: absoluteFile(template.path, "template.path"), sha256: digest(template.sha256, "template.sha256") } : undefined,
    designMode: designMode as "shared" | "template",
    cases: normalizedCases,
  };
}

export async function qualifyPrivateGolden(manifestPath: string, outputRoot: string) {
  const manifestBytes = new Uint8Array(await fs.readFile(manifestPath));
  const manifest = parsePrivateGoldenManifest(JSON.parse(new TextDecoder().decode(manifestBytes)));
  const manifestSha256 = await sha256(manifestBytes);
  if (manifest.template) {
    const templateSha256 = await sha256(new Uint8Array(await fs.readFile(manifest.template.path)));
    if (templateSha256 !== manifest.template.sha256) throw new Error(`The authorized template hash changed. Expected ${manifest.template.sha256}, received ${templateSha256}.`);
  }
  const benchmarkSlideNumbers = [...new Set(manifest.cases.flatMap((item) => item.benchmarkSlides))];
  const result = await qualifyStudioWebBenchmark(
    manifest.source.path,
    manifest.cases.map((item) => item.sourceSlide),
    outputRoot,
    manifest.benchmark.path,
    { sourceSha256: manifest.source.sha256, benchmarkSha256: manifest.benchmark.sha256 },
    { templatePath: manifest.template?.path, designMode: manifest.designMode, benchmarkSlideNumbers },
  );
  const impactBySlide = new Map(result.report.designImpact.map((impact) => [impact.sourceSlideNumber, impact]));
  const cases = manifest.cases.map((item) => ({
    ...item,
    designImpact: impactBySlide.get(item.sourceSlide),
    evidence: {
      source: path.join(outputRoot, `source-slide-${String(item.sourceSlide).padStart(2, "0")}.png`),
      candidate: path.join(outputRoot, `candidate-slide-${String(item.sourceSlide).padStart(2, "0")}.png`),
      benchmark: item.benchmarkSlides.map((slide) => path.join(outputRoot, `visual-benchmark-slide-${String(slide).padStart(2, "0")}.png`)),
    },
    visualAcceptance: {
      status: "pending-human-or-authorized-ai-review" as const,
      requiredJudgments: item.reviewFocus,
      rule: "Inspect the source, exact candidate, and approved benchmark at full-slide size. Objective checks make the candidate reviewable; they do not prove it is better.",
    },
  }));
  const ledger = {
    schema: "presentation-studio/private-golden-qualification",
    version: 2,
    generatedAt: new Date().toISOString(),
    manifest: { id: manifest.id, path: manifestPath, sha256: manifestSha256 },
    source: result.report.source,
    benchmark: result.report.benchmark,
    candidate: { path: result.report.candidate.path, sha256: result.report.candidate.sha256 },
    objectiveChecks: result.report.checks,
    objectiveReady: result.report.readyForHumanVisualReview,
    status: result.report.readyForHumanVisualReview ? "visual-review-required" : "objective-failure",
    cases,
    acceptanceRule: "Do not mark this golden qualification passed until every case receives an evidence-backed visual judgment and the exact candidate PPTX passes editability and native-export review.",
  };
  const ledgerPath = path.join(outputRoot, "private-golden-qualification.json");
  await fs.writeFile(ledgerPath, `${JSON.stringify(ledger, null, 2)}\n`);
  const escapeHtml = (value: unknown) => String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
  const relativeImage = (imagePath: string) => escapeHtml(path.relative(outputRoot, imagePath));
  const caseHtml = cases.map((item) => `
    <section class="case">
      <header><div><strong>${escapeHtml(item.id)}</strong><span>Source slide ${item.sourceSlide}</span></div><p>${escapeHtml(item.communicationJob)}</p></header>
      <div class="triptych">
        <figure><figcaption>Immutable source</figcaption><img src="${relativeImage(item.evidence.source)}" alt="Source slide ${item.sourceSlide}"></figure>
        <figure><figcaption>Studio candidate</figcaption><img src="${relativeImage(item.evidence.candidate)}" alt="Studio candidate for source slide ${item.sourceSlide}"></figure>
        ${item.evidence.benchmark.map((imagePath, index) => `<figure><figcaption>Human-cleaned golden${item.evidence.benchmark.length > 1 ? ` ${index + 1}` : ""}</figcaption><img src="${relativeImage(imagePath)}" alt="Golden benchmark"></figure>`).join("")}
      </div>
      <div class="review"><b>Required review:</b> ${item.reviewFocus.map(escapeHtml).join(" · ")}<br><b>Recorded design impact:</b> ${escapeHtml(item.designImpact?.summary ?? "unrecorded")}</div>
    </section>`).join("");
  const htmlPath = path.join(outputRoot, "private-golden-review.html");
  await fs.writeFile(htmlPath, `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Presentation Studio private golden review</title><style>body{margin:0;background:#eef1ef;color:#25312b;font:15px Aptos,Arial,sans-serif}main{max-width:1800px;margin:auto;padding:28px}.summary,.case{background:#fff;border:1px solid #d7ded9;margin:0 0 24px;padding:20px}.summary{border-top:6px solid #007833}.case header{display:flex;gap:30px;align-items:flex-start;justify-content:space-between}.case header div{display:flex;gap:12px;align-items:baseline}.case header span{color:#607068}.case header p{max-width:820px;margin:0}.triptych{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px;margin-top:18px}figure{margin:0;border:1px solid #d7ded9;background:#f7f8f7}figcaption{padding:9px 12px;background:#25312b;color:#fff;font-weight:700}img{display:block;width:100%;height:auto}.review{margin-top:14px;padding:12px;background:#f2f5f3;line-height:1.6}.fail{color:#9c2b19}.pass{color:#007833}</style></head><body><main><section class="summary"><h1>Private golden visual qualification</h1><p><b>Manifest:</b> ${escapeHtml(manifest.id)} · <b>Objective status:</b> <span class="${ledger.objectiveReady ? "pass" : "fail"}">${escapeHtml(ledger.status)}</span></p><p>This page compares native PowerPoint pixels. It is deliberately not an automatic aesthetic pass: every case still requires an evidence-backed source/candidate/golden judgment.</p></section>${caseHtml}</main></body></html>`);
  const promptPath = path.join(outputRoot, "fresh-agent-prompt.txt");
  await fs.writeFile(promptPath, `Connect to the Presentation Studio MCP and work only in the currently open project. Read get_design_contract, get_app_status, get_agent_runbook, and the deck work order. Follow the runbook one action at a time. Prove the representative archetype set before propagating any design pattern deck-wide. For each representative, inspect the immutable source pixels, build the exact Studio candidate in native PowerPoint, inspect both full-size images, and apply the source-wins rule. Preserve exact content and meaning-bearing visual relationships. ORNL is the default. Do not save or export. Report product or engine defects precisely rather than working around them.\n`);
  return { ledger, ledgerPath, htmlPath, promptPath, benchmarkReportPath: result.reportPath };
}

if (path.resolve(fileURLToPath(import.meta.url)) === path.resolve(process.argv[1] ?? "")) {
  const manifestPath = argument("manifest");
  if (!manifestPath) throw new Error("Use --manifest /absolute/path/to/private-golden-manifest.json.");
  const outputRoot = argument("output") ?? path.join(os.tmpdir(), "presentation-studio-private-golden");
  qualifyPrivateGolden(manifestPath, outputRoot).then(({ ledger, ledgerPath, htmlPath, promptPath, benchmarkReportPath }) => {
    process.stdout.write(`${JSON.stringify({ status: ledger.status, objectiveReady: ledger.objectiveReady, candidate: ledger.candidate.path, ledger: ledgerPath, review: htmlPath, freshAgentPrompt: promptPath, benchmarkReport: benchmarkReportPath }, null, 2)}\n`);
    if (!ledger.objectiveReady) process.exitCode = 1;
  }).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
