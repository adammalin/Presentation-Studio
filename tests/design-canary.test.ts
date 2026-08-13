import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createDesignCanaryDeck } from "../scripts/create-design-canary";
import { runDesignCanaryCli } from "../scripts/run-design-canary";
import { auditPptx } from "../src/lib/pptx-audit";
import { compilePresentationScene } from "../src/lib/scene-graph";
import type { DeckJob } from "../src/types";

interface CanaryGroundTruth {
  schema: string;
  version: number;
  slideSize: string;
  fontFamily: string;
  slides: Array<{
    slideNumber: number;
    defect: string;
    objects: string[];
    expectedDetection: boolean;
    requiresAiJudgment?: boolean;
  }>;
}

test("precision-layout canary generator matches its complete versioned ground truth", async () => {
  const root = path.resolve(import.meta.dirname, "..");
  const groundTruth = JSON.parse(await fs.readFile(path.join(root, "fixtures", "design-canary-ground-truth.json"), "utf8")) as CanaryGroundTruth;
  assert.equal(groundTruth.schema, "presentation-studio/design-canary-ground-truth");
  assert.equal(groundTruth.version, 2);
  assert.equal(groundTruth.slideSize, "16:9");
  assert.equal(groundTruth.fontFamily, "Aptos");
  assert.deepEqual(groundTruth.slides.map((slide) => slide.slideNumber), Array.from({ length: 14 }, (_value, index) => index + 1));
  assert.equal(new Set(groundTruth.slides.flatMap((slide) => slide.objects)).size, groundTruth.slides.flatMap((slide) => slide.objects).length);
  assert.equal(groundTruth.slides.filter((slide) => slide.requiresAiJudgment).length, 1);
  assert.equal(groundTruth.slides.find((slide) => slide.requiresAiJudgment)?.slideNumber, 14);

  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-canary-contract-"));
  const filePath = path.join(directory, "precision-layout-canary.pptx");
  await createDesignCanaryDeck(filePath);
  const bytes = new Uint8Array(await fs.readFile(filePath));
  const audit = await auditPptx(bytes);
  assert.equal(audit.slideCount, 14);
  assert.ok(Math.abs(audit.slideSize.width / audit.slideSize.height - 16 / 9) < .001);

  const deck: DeckJob = {
    id: "canary-contract",
    name: path.basename(filePath),
    sourceResourceId: "canary-source",
    sourceSha256: "c".repeat(64),
    operationScope: "reflow",
    templateClassification: "custom",
    targetTemplateId: "synthetic-canary",
    targetTemplateConfirmedAt: "2026-08-12T20:00:00.000Z",
    status: "ready-for-cleanup",
    audit,
    protectedSlideNumbers: [],
  };
  const scene = compilePresentationScene({ ...deck, audit });
  const objectByName = new Map(scene.objects.map((object) => [object.name, object]));
  for (const expectation of groundTruth.slides) {
    assert.ok(expectation.defect.trim());
    assert.ok(expectation.objects.length > 0);
    for (const name of expectation.objects) {
      const object = objectByName.get(name);
      assert.ok(object, `Generated canary is missing ${name}.`);
      assert.equal(object.slideNumber, expectation.slideNumber, `${name} is bound to the wrong slide.`);
    }
  }

  for (const slideNumber of [7, 8, 9, 10, 11, 12, 13]) {
    assert.equal(audit.tables.filter((table) => table.slideNumber === slideNumber).length, 1, `Slide ${slideNumber} must carry one native table.`);
  }
  const mergedTable = audit.tables.find((table) => table.slideNumber === 13);
  assert.ok(mergedTable && mergedTable.mergedCellCount > 0, "The merged-cell canary must retain native merge topology.");
});

test("precision-layout canary CLI returns failure and persists evidence when native validation cannot run", async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-canary-cli-"));
  const stdout: string[] = [];
  const stderr: string[] = [];
  const exitCode = await runDesignCanaryCli(root, {
    run: async () => {
      throw new Error("PowerPoint-native measurement unavailable for regression test.");
    },
    stdout: (message) => stdout.push(message),
    stderr: (message) => stderr.push(message),
  });

  assert.equal(exitCode, 1);
  assert.deepEqual(stdout, []);
  assert.equal(stderr.length, 1);
  const report = JSON.parse(await fs.readFile(path.join(root, "tmp", "design-canary", "report.json"), "utf8"));
  assert.equal(report.passed, false);
  assert.equal(report.checks.canaryCompleted, false);
  assert.match(report.error, /PowerPoint-native measurement unavailable/);
});
