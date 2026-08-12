import assert from "node:assert/strict";
import test from "node:test";
import packagedStandard from "../shared/presentation-design-standard.json";
import { PRESENTATION_DESIGN_STANDARD } from "../src/lib/design-standard";
import { createProject, projectSchema } from "../src/lib/project";

test("one versioned design standard drives project defaults and MCP", () => {
  const project = createProject("Design standard test");
  assert.equal(PRESENTATION_DESIGN_STANDARD.version, packagedStandard.version);
  assert.equal(project.settings.designStandardVersion, packagedStandard.version);
  assert.equal(project.settings.defaultSlideSize, "16:9");
  assert.equal(project.settings.defaultFontFamily, "Aptos");
  assert.equal(packagedStandard.tableProfile.fontFamily, "Aptos");
  assert.equal(packagedStandard.tableProfile.cellPaddingPt.left, 6);
  assert.equal(packagedStandard.tableProfile.strokes.outer, "none");
  assert.match(packagedStandard.autonomy.approvalPolicy, /routine deterministic design choices/i);
});

test("older project metadata adopts design defaults without inventing review threads", () => {
  const legacy = createProject("Legacy test") as unknown as Record<string, unknown>;
  const settings = { ...(legacy.settings as Record<string, unknown>) };
  delete settings.designStandardVersion;
  delete settings.defaultProfileId;
  delete settings.defaultSlideSize;
  delete settings.defaultFontFamily;
  legacy.settings = settings;
  delete legacy.designThreads;
  const parsed = projectSchema.parse(legacy);
  assert.equal(parsed.settings.defaultSlideSize, "16:9");
  assert.equal(parsed.settings.defaultFontFamily, "Aptos");
  assert.deepEqual(parsed.designThreads, []);
});
