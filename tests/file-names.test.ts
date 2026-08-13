import assert from "node:assert/strict";
import test from "node:test";
import { cleanFileStem, projectSaveDefaultName } from "../src/lib/file-names";

test("project saves inherit the source PowerPoint name", () => {
  assert.equal(projectSaveDefaultName({
    projectName: "Untitled review batch",
    deckNames: ["5_System_level_EMT_Simulations.pptx"],
    encrypted: false,
  }), "5_System_level_EMT_Simulations_Presentation-Studio.pstudio");

  assert.equal(projectSaveDefaultName({
    projectName: "Untitled review batch",
    deckNames: ["5_System_level_EMT_Simulations.pptx"],
    encrypted: true,
  }), "5_System_level_EMT_Simulations_Presentation-Studio.pstudio-secure");
});

test("multi-deck project names identify the first source and remaining deck count", () => {
  assert.equal(projectSaveDefaultName({
    projectName: "Review batch",
    deckNames: ["Primary deck.pptx", "Sponsor deck.pptx", "Appendix.pptx"],
    encrypted: false,
  }), "Primary deck_and-2-more_Presentation-Studio.pstudio");
});

test("project saves fall back to the project title before a PowerPoint is imported", () => {
  assert.equal(projectSaveDefaultName({
    projectName: "August review batch",
    deckNames: [],
    encrypted: false,
  }), "August review batch.pstudio");
});

test("file stems remove presentation extensions and unsafe path characters", () => {
  assert.equal(cleanFileStem("Sponsor/Study:Final.PPTM"), "Sponsor-Study-Final");
});
