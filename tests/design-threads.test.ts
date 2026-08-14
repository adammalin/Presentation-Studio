import assert from "node:assert/strict";
import test from "node:test";
import { markSubmittedThreadsForReanchor, removeAddressedDesignThreads, removeAddressedDesignThreadsForSlides, removeCompletedDesignThreads, removeDesignThread } from "../src/lib/design-threads";
import type { CleanupProposal, DesignThread } from "../src/types";

test("staging shared components removes addressed submitted comments from the clean slide", () => {
  const threads: DesignThread[] = [
    { id: "addressed", deckId: "deck-a", slideId: "slide-4", slideNumber: 4, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Fix spacing", status: "submitted", createdAt: "r1", updatedAt: "r1" },
    { id: "unaddressed", deckId: "deck-a", slideId: "slide-5", slideNumber: 5, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Fix image", status: "submitted", createdAt: "r1", updatedAt: "r1" },
    { id: "private-note", deckId: "deck-a", slideId: "slide-4", slideNumber: 4, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Private", status: "note", createdAt: "r1", updatedAt: "r1" },
  ];
  const proposal = { changes: [{ id: "component-slide-4", kind: "text-style", selected: true, affectedSlideNumbers: [4] }] } as CleanupProposal;
  const result = removeAddressedDesignThreads(threads, "deck-a", proposal, ["addressed"]);
  assert.deepEqual(result.map((thread) => thread.id), ["unaddressed", "private-note"]);
  assert.equal(removeAddressedDesignThreads(threads, "deck-a", proposal).length, 3);
});

test("legacy completed comments are pruned while active comments and private notes remain", () => {
  const threads: DesignThread[] = [
    { id: "submitted", deckId: "deck-a", slideId: "slide-1", slideNumber: 1, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Active", status: "submitted", createdAt: "r1", updatedAt: "r1" },
    { id: "ready", deckId: "deck-a", slideId: "slide-2", slideNumber: 2, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Done", status: "proposal-ready", createdAt: "r1", updatedAt: "r1" },
    { id: "resolved", deckId: "deck-a", slideId: "slide-3", slideNumber: 3, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Done", status: "resolved", createdAt: "r1", updatedAt: "r1" },
    { id: "note", deckId: "deck-a", slideId: "slide-4", slideNumber: 4, baseRevision: "r1", anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment: "Private", status: "note", createdAt: "r1", updatedAt: "r1" },
  ];
  assert.deepEqual(removeCompletedDesignThreads(threads).map((thread) => thread.id), ["submitted", "note"]);
  assert.deepEqual(removeDesignThread(threads, "submitted").map((thread) => thread.id), ["ready", "resolved", "note"]);
});

test("a changed Studio revision marks unaddressed submitted comments for reanchor and clears explicitly addressed ones", () => {
  const threads: DesignThread[] = [
    { id: "addressed", deckId: "deck-a", slideId: "slide-4", slideNumber: 4, baseRevision: "r1", anchor: { kind: "region", x: .1, y: .1, width: .2, height: .2 }, comment: "Fix spacing", status: "submitted", createdAt: "r1", updatedAt: "r1" },
    { id: "still-open", deckId: "deck-a", slideId: "slide-4", slideNumber: 4, baseRevision: "r1", anchor: { kind: "region", x: .6, y: .1, width: .2, height: .2 }, comment: "Also check this", status: "submitted", createdAt: "r1", updatedAt: "r1" },
    { id: "private", deckId: "deck-a", slideId: "slide-4", slideNumber: 4, baseRevision: "r1", anchor: { kind: "region", x: .1, y: .6, width: .2, height: .2 }, comment: "Private", status: "note", createdAt: "r1", updatedAt: "r1" },
  ];
  const rebound = markSubmittedThreadsForReanchor(threads, "deck-a", 4, "r2", ["addressed"]);
  assert.equal(rebound.find((thread) => thread.id === "still-open")?.status, "needs-reanchor");
  assert.equal(rebound.find((thread) => thread.id === "private")?.status, "note");
  const result = removeAddressedDesignThreadsForSlides(rebound, "deck-a", [4], ["addressed"]);
  assert.deepEqual(result.map((thread) => thread.id), ["still-open", "private"]);
});
