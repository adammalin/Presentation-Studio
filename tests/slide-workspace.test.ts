import assert from "node:assert/strict";
import test from "node:test";
import { isProposalSlideWorkspaceRequest, type SlideWorkspaceRequest } from "../src/lib/slide-workspace";

test("resources-only projects do not become proposal workspaces when both IDs are absent", () => {
  assert.equal(isProposalSlideWorkspaceRequest(undefined, undefined), false);
  assert.equal(isProposalSlideWorkspaceRequest(undefined, "deck-1"), false);
});

test("proposal workspace matching requires one real request for the selected deck", () => {
  const request: SlideWorkspaceRequest = { id: "request-1", deckId: "deck-1", slideNumber: 4, mode: "review", representation: "proposal" };
  assert.equal(isProposalSlideWorkspaceRequest(request, "deck-1"), true);
  assert.equal(isProposalSlideWorkspaceRequest(request, "deck-2"), false);
  assert.equal(isProposalSlideWorkspaceRequest({ ...request, representation: "current" }, "deck-1"), false);
});
