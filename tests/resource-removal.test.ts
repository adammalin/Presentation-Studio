import assert from "node:assert/strict";
import test from "node:test";
import type { DeckJob, ProjectResource } from "../src/types";
import { createProject } from "../src/lib/project";
import { removeResourceFromProject, resourceRemovalImpact } from "../src/lib/resource-removal";

function resource(id: string, name: string): ProjectResource {
  return { id, name, mediaType: "application/octet-stream", byteLength: 1, sha256: id.padEnd(64, "0").slice(0, 64), roles: ["grounding-source"], createdAt: "2026-08-13T12:00:00.000Z", embedded: true, bytes: new Uint8Array([1]), mcpAccess: "none" };
}

test("removing an ordinary Resource changes only the self-contained project", () => {
  const project = createProject("Removal test");
  project.resources = [resource("a", "notes.txt"), resource("b", "image.png")];
  const result = removeResourceFromProject(project, "a");
  assert.deepEqual(result.project.resources.map((item) => item.id), ["b"]);
  assert.equal(result.project.decks.length, 0);
  assert.equal(project.resources.length, 2);
  assert.match(result.project.activity.at(-1)?.detail ?? "", /No external source file was changed or deleted/);
});

test("removing a deck source also removes dependent project state without touching unrelated work", () => {
  const project = createProject("Linked removal test");
  project.resources = [resource("deck-source", "source.pptx"), resource("notes", "notes.txt")];
  const linkedDeck: DeckJob = { id: "deck-one", name: "source.pptx", sourceResourceId: "deck-source", sourceSha256: "d".repeat(64), operationScope: "reflow", templateClassification: "custom", status: "audited", protectedSlideNumbers: [] };
  const otherDeck: DeckJob = { ...linkedDeck, id: "deck-two", name: "other.pptx", sourceResourceId: "notes" };
  project.decks = [linkedDeck, otherDeck];
  project.styleExemplars = [{ id: "example", name: "Table", kind: "table", resourceId: "deck-source", deckId: "deck-one", slideNumber: 1, objectOrdinal: 1, scope: "deck", createdAt: "2026-08-13T12:00:00.000Z" }];
  project.designThreads = [{ id: "thread", deckId: "deck-one", slideId: "slide-1", slideNumber: 1, baseRevision: "2026-08-13T12:00:00.000Z", anchor: { kind: "region", x: 0, y: 0, width: .1, height: .1 }, comment: "Adjust table", status: "submitted", createdAt: "2026-08-13T12:00:00.000Z", updatedAt: "2026-08-13T12:00:00.000Z" }];
  const impact = resourceRemovalImpact(project, "deck-source");
  assert.deepEqual(impact.linkedDeckIds, ["deck-one"]);
  assert.equal(impact.removedExemplarCount, 1);
  assert.equal(impact.removedThreadCount, 1);
  const result = removeResourceFromProject(project, "deck-source");
  assert.deepEqual(result.project.resources.map((item) => item.id), ["notes"]);
  assert.deepEqual(result.project.decks.map((deck) => deck.id), ["deck-two"]);
  assert.equal(result.project.styleExemplars.length, 0);
  assert.equal(result.project.designThreads.length, 0);
});
