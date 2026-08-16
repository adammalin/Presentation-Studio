import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { auditPptx } from "../src/lib/pptx-audit";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { bindNativeMeasurement } from "../src/lib/native-measurement";
import { calculateDesignMetrics } from "../src/lib/design-metrics";
import { buildDeckQualificationReport } from "../src/lib/deck-qualification";
import { sha256 } from "../src/lib/hash";
import type { DeckJob } from "../src/types";

const require = createRequire(import.meta.url);
const { captureDeckQualification, finalizeDeckQualification } = require("../electron/deck-qualification.cjs") as {
  captureDeckQualification(input: { source: { name: string; bytes: Uint8Array }; candidate: { name: string; bytes: Uint8Array }; outputRoot: string; width: number }): Promise<{
    outputRoot: string;
    sourceRender: Parameters<typeof buildDeckQualificationReport>[0]["sourceRender"];
    candidateRender: Parameters<typeof buildDeckQualificationReport>[0]["candidateRender"];
    sourceMeasurement: Parameters<typeof buildDeckQualificationReport>[0]["sourceMeasurement"];
    candidateMeasurement: Parameters<typeof buildDeckQualificationReport>[0]["candidateMeasurement"];
  }>;
  finalizeDeckQualification(input: { outputRoot: string; report: ReturnType<typeof buildDeckQualificationReport> }): Promise<{ outputRoot: string; reportPath: string; htmlPath: string }>;
};

function argument(name: string) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function absolutePptx(value: string | undefined, label: string) {
  if (!value || !path.isAbsolute(value) || !/\.pptx$/i.test(value)) throw new Error(`${label} must be an absolute PPTX path.`);
  return path.resolve(value);
}

function protectedSlides(value: string | undefined) {
  if (!value) return [];
  const slides = value.split(",").map((item) => Number(item.trim())).filter((item) => Number.isInteger(item) && item > 0);
  if (!slides.length || slides.length !== new Set(slides).size) throw new Error("--protected-slides must contain unique positive slide numbers separated by commas.");
  return slides;
}

async function candidateDeck(name: string, bytes: Uint8Array, audit: Awaited<ReturnType<typeof auditPptx>>): Promise<DeckJob> {
  const digest = await sha256(bytes);
  const deck: DeckJob = {
    id: "qualification-candidate",
    name,
    sourceResourceId: "qualification-candidate-resource",
    sourceSha256: digest,
    operationScope: "reflow",
    templateClassification: audit.classification,
    status: "ready-for-cleanup",
    audit,
    protectedSlideNumbers: [],
  };
  deck.scene = compilePresentationScene({ ...deck, audit });
  return deck;
}

export async function qualifyDeck(sourcePath: string, candidatePath: string, outputRoot: string, options: { width?: number; protectedSlideNumbers?: number[] } = {}) {
  const [sourceBuffer, candidateBuffer] = await Promise.all([fs.readFile(sourcePath), fs.readFile(candidatePath)]);
  const sourceBytes = new Uint8Array(sourceBuffer);
  const candidateBytes = new Uint8Array(candidateBuffer);
  const [sourceAudit, candidateAudit, sourceSha256, candidateSha256] = await Promise.all([
    auditPptx(sourceBytes),
    auditPptx(candidateBytes),
    sha256(sourceBytes),
    sha256(candidateBytes),
  ]);
  const capture = await captureDeckQualification({
    source: { name: path.basename(sourcePath), bytes: sourceBytes },
    candidate: { name: path.basename(candidatePath), bytes: candidateBytes },
    outputRoot,
    width: options.width ?? 2200,
  });
  const deck = await candidateDeck(path.basename(candidatePath), candidateBytes, candidateAudit);
  const candidateMeasurement = bindNativeMeasurement(deck, capture.candidateMeasurement);
  const candidateMetrics = calculateDesignMetrics(deck, candidateMeasurement);
  const report = buildDeckQualificationReport({
    id: path.basename(outputRoot),
    sourceName: path.basename(sourcePath),
    candidateName: path.basename(candidatePath),
    sourceSha256,
    candidateSha256,
    sourceAudit,
    candidateAudit,
    sourceRender: capture.sourceRender,
    candidateRender: capture.candidateRender,
    sourceMeasurement: capture.sourceMeasurement,
    candidateMeasurement: capture.candidateMeasurement,
    candidateMeasurementPacket: candidateMeasurement,
    candidateMetrics,
    protectedSlideNumbers: options.protectedSlideNumbers,
  });
  const finalized = await finalizeDeckQualification({ outputRoot: capture.outputRoot, report });
  return { report, ...finalized };
}

if (path.resolve(fileURLToPath(import.meta.url)) === path.resolve(process.argv[1] ?? "")) {
  const sourcePath = absolutePptx(argument("source"), "--source");
  const candidatePath = absolutePptx(argument("candidate"), "--candidate");
  const outputRoot = path.resolve(argument("output") ?? path.join(os.tmpdir(), `presentation-studio-qualification-${Date.now()}`));
  const width = Number(argument("width") ?? 2200);
  qualifyDeck(sourcePath, candidatePath, outputRoot, { width, protectedSlideNumbers: protectedSlides(argument("protected-slides")) }).then(({ report, reportPath, htmlPath }) => {
    process.stdout.write(`${JSON.stringify({ status: report.status, slides: report.totals.slides, blockers: report.totals.blockerIssues, majorIssues: report.totals.majorIssues, report: reportPath, review: htmlPath }, null, 2)}\n`);
    if (report.status === "objective-failure") process.exitCode = 1;
  }).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
