import assert from "node:assert/strict";
import test from "node:test";
import { isPresentationStudioProjectName, projectPackageFromDrop } from "../src/lib/project-drop";

test("a single Presentation Studio package is routed to project opening", () => {
  const project = { name: "review.pstudio", marker: 1 };
  assert.equal(isPresentationStudioProjectName(project.name), true);
  assert.equal(isPresentationStudioProjectName("review.PSTUDIO-SECURE"), true);
  assert.equal(projectPackageFromDrop([project]), project);
  assert.equal(projectPackageFromDrop([{ name: "source.pptx" }]), undefined);
});

test("project drops reject ambiguous mixed batches", () => {
  assert.throws(() => projectPackageFromDrop([{ name: "review.pstudio" }, { name: "notes.docx" }]), /one \.pstudio project/i);
  assert.throws(() => projectPackageFromDrop([{ name: "one.pstudio" }, { name: "two.pstudio" }]), /one \.pstudio project/i);
});
