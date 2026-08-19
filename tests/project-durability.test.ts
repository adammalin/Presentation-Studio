import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { sha256 } from "../src/lib/hash";
import { auditPptx } from "../src/lib/pptx-audit";
import { createProject, touchProject } from "../src/lib/project";
import { buildProjectPackage, openProjectPackage } from "../src/lib/project-package";
import { applyProjectRecoveryCheckpoint, buildProjectRecoveryCheckpoint, createLatestOnlySaver, projectHasRecoverableWork, type AutosaveProgress } from "../src/lib/project-durability";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { buildSlideRenderCatalog } from "../src/lib/template-catalog";
import { compileStudioWebScene, recomposeStudioWebSlide } from "../src/lib/studio-web-scene";
import type { DeckJob, PresentationStudioProject } from "../src/types";

async function projectWithStudioScene(): Promise<{ base: PresentationStudioProject; designed: PresentationStudioProject }> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-recovery-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const bytes = new Uint8Array(await fs.readFile(filePath));
  const digest = await sha256(bytes);
  const audit = await auditPptx(bytes);
  const project = createProject("Recovery test");
  project.resources.push({
    id: "source",
    name: "synthetic.pptx",
    mediaType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    byteLength: bytes.byteLength,
    sha256: digest,
    roles: ["grounding-source"],
    kind: "presentation",
    support: ["source-readable", "previewable"],
    processing: { status: "indexed", summary: "Synthetic presentation indexed.", processedAt: new Date().toISOString(), warnings: [] },
    createdAt: new Date().toISOString(),
    embedded: true,
    bytes,
    mcpAccess: "none",
  });
  const deck: DeckJob = { id: "deck", name: "synthetic.pptx", sourceResourceId: "source", sourceSha256: digest, operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  project.decks.push(deck);
  const base = structuredClone(project);
  base.resources[0].bytes = bytes;
  const catalog = await buildSlideRenderCatalog(bytes, deck.name);
  const studioScene = recomposeStudioWebSlide(compileStudioWebScene(deck, catalog), 1, "ornl-title-content");
  const designed = touchProject({ ...project, decks: [{ ...deck, studioScene }] }, "studio-slide-designed", "Designed slide 1 in the central Studio scene.");
  return { base, designed };
}

test("a new untouched workspace does not overwrite the prior recoverable project", () => {
  const blank = createProject();
  assert.equal(projectHasRecoverableWork(blank), false);
  assert.equal(projectHasRecoverableWork(touchProject(blank, "project-renamed", "Named the new project.")), true);
});

test("a lightweight recovery checkpoint restores the latest Studio scene over matching embedded Resources", async () => {
  const { base, designed } = await projectWithStudioScene();
  const completeAutosave = await openProjectPackage(await buildProjectPackage(base));
  const recovered = applyProjectRecoveryCheckpoint(completeAutosave, buildProjectRecoveryCheckpoint(designed));
  assert.equal(recovered.project.updatedAt, designed.project.updatedAt);
  assert.equal(recovered.decks[0].studioScene?.revision, designed.decks[0].studioScene?.revision);
  assert.equal(recovered.decks[0].studioScene?.slides[0].status, "designed");
  assert.deepEqual(recovered.resources[0].bytes, base.resources[0].bytes);
  assert.equal(recovered.resources[0].mcpAccess, "none");
});

test("a recovery checkpoint cannot attach to a different project or Resource inventory", async () => {
  const { base, designed } = await projectWithStudioScene();
  const checkpoint = buildProjectRecoveryCheckpoint(designed);
  assert.throws(() => applyProjectRecoveryCheckpoint({ ...base, project: { ...base.project, id: "different-project" } }, checkpoint), /different project/i);
  assert.throws(() => applyProjectRecoveryCheckpoint({ ...base, resources: [] }, checkpoint), /does not match/i);
});

test("latest-only autosave coalesces a burst and durably writes the newest revision", async () => {
  const saved: number[] = [];
  const progress: AutosaveProgress[] = [];
  const saver = createLatestOnlySaver<number>({ delayMs: 50, revision: (value) => String(value), save: async (value) => { saved.push(value); }, onProgress: (value) => progress.push(value) });
  saver.request(1);
  saver.request(2);
  saver.request(3);
  await saver.flush();
  assert.deepEqual(saved, [3]);
  assert.equal(progress.at(-1)?.phase, "saved");
  assert.equal(progress.at(-1)?.revision, "3");
  await saver.dispose();
});

test("latest-only autosave surfaces a write failure instead of silently discarding it", async () => {
  const progress: AutosaveProgress[] = [];
  const saver = createLatestOnlySaver<number>({ delayMs: 0, revision: String, save: async () => { throw new Error("disk unavailable"); }, onProgress: (value) => progress.push(value) });
  saver.request(1);
  await assert.rejects(() => saver.flush(), /disk unavailable/);
  assert.equal(progress.at(-1)?.phase, "error");
  await saver.dispose().catch(() => undefined);
});
