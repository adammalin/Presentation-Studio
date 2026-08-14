import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { sha256 } from "../src/lib/hash";
import { qualifyStudioWebBenchmark } from "./qualify-studio-web-benchmark";

interface PrivateGoldenManifest {
  schema: "presentation-studio/private-golden-manifest";
  version: 1;
  id: string;
  source: { path: string; sha256: string };
  benchmark: { path: string; sha256: string };
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
  if (raw.schema !== "presentation-studio/private-golden-manifest" || raw.version !== 1) throw new Error("Use presentation-studio/private-golden-manifest version 1.");
  if (typeof raw.id !== "string" || !/^[a-z0-9._-]{1,120}$/i.test(raw.id)) throw new Error("The private golden manifest requires a stable ID.");
  const source = raw.source as Record<string, unknown>;
  const benchmark = raw.benchmark as Record<string, unknown>;
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
    version: 1,
    id: raw.id,
    source: { path: absoluteFile(source.path, "source.path"), sha256: digest(source.sha256, "source.sha256") },
    benchmark: { path: absoluteFile(benchmark.path, "benchmark.path"), sha256: digest(benchmark.sha256, "benchmark.sha256") },
    cases: normalizedCases,
  };
}

export async function qualifyPrivateGolden(manifestPath: string, outputRoot: string) {
  const manifestBytes = new Uint8Array(await fs.readFile(manifestPath));
  const manifest = parsePrivateGoldenManifest(JSON.parse(new TextDecoder().decode(manifestBytes)));
  const manifestSha256 = await sha256(manifestBytes);
  const result = await qualifyStudioWebBenchmark(
    manifest.source.path,
    manifest.cases.map((item) => item.sourceSlide),
    outputRoot,
    manifest.benchmark.path,
    { sourceSha256: manifest.source.sha256, benchmarkSha256: manifest.benchmark.sha256 },
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
    version: 1,
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
  return { ledger, ledgerPath, benchmarkReportPath: result.reportPath };
}

if (path.resolve(fileURLToPath(import.meta.url)) === path.resolve(process.argv[1] ?? "")) {
  const manifestPath = argument("manifest");
  if (!manifestPath) throw new Error("Use --manifest /absolute/path/to/private-golden-manifest.json.");
  const outputRoot = argument("output") ?? path.join(os.tmpdir(), "presentation-studio-private-golden");
  qualifyPrivateGolden(manifestPath, outputRoot).then(({ ledger, ledgerPath, benchmarkReportPath }) => {
    process.stdout.write(`${JSON.stringify({ status: ledger.status, objectiveReady: ledger.objectiveReady, candidate: ledger.candidate.path, ledger: ledgerPath, benchmarkReport: benchmarkReportPath }, null, 2)}\n`);
    if (!ledger.objectiveReady) process.exitCode = 1;
  }).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
