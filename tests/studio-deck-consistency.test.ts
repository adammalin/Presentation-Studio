import assert from "node:assert/strict";
import test from "node:test";
import { analyzeStudioDeckConsistency } from "../src/lib/studio-deck-consistency";
import { STUDIO_WEB_SCENE_SCHEMA, STUDIO_WEB_SCENE_VERSION, type StudioWebNode, type StudioWebScene } from "../src/types";

const emu = (inches: number) => inches * 914_400;

function node(id: string, role: StudioWebNode["role"], slideOffset = 0): StudioWebNode {
  const frame = { x: emu(.5 + slideOffset), y: emu(.5), width: emu(4), height: emu(.5), rotation: 0 };
  return { id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "semantic-atom", name: id, kind: "text", role, sourceFrame: frame, frame, zIndex: 1, sourceTextOrder: 1, visible: true, locked: false, exactContent: true, text: id, textHash: "a".repeat(64), component: role === "title" ? undefined : { groupId: `group-${id}`, role: "card-body" }, style: { fontFamily: "Aptos", fontSizePt: 16, fontWeight: 400, lineHeight: 1.08, color: "#373A36", borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 4, right: 7, bottom: 4, left: 7 } } };
}

function tableNode(id: string, borderMode: "none" | "subtle"): StudioWebNode {
  const value = node(id, "table");
  return { ...value, kind: "table", text: undefined, textHash: undefined, tableId: id, table: { rows: 2, columns: 2, cells: [{ id: `${id}-a`, row: 1, column: 1, rowSpan: 1, columnSpan: 1, text: "A" }, { id: `${id}-b`, row: 1, column: 2, rowSpan: 1, columnSpan: 1, text: "B" }, { id: `${id}-c`, row: 2, column: 1, rowSpan: 1, columnSpan: 1, text: "C" }, { id: `${id}-d`, row: 2, column: 2, rowSpan: 1, columnSpan: 1, text: "D" }], design: { headerRows: 1, columnWidths: [.5, .5], rowHeights: [.5, .5], borderMode, borderColor: "#DBDCDB", borderWidthPt: borderMode === "none" ? 0 : .75, defaultPaddingPt: { top: 4, right: 7, bottom: 4, left: 7 }, cellStyles: [] } } };
}

test("deck consistency review identifies title, repeated-component, and related-table outliers", () => {
  const titles = [node("title-1", "title"), node("title-2", "title"), node("title-3", "title", .35)];
  const bodies = [node("body-1", "body"), node("body-2", "body"), { ...node("body-3", "body"), style: { ...node("body-3", "body").style, fontSizePt: 13 } }];
  const tables = [tableNode("table-1", "subtle"), tableNode("table-2", "none")];
  const scene: StudioWebScene = { schema: STUDIO_WEB_SCENE_SCHEMA, version: STUDIO_WEB_SCENE_VERSION, revision: "scene-revision", deckId: "deck", sourceSha256: "b".repeat(64), slideSize: { width: emu(13.333), height: emu(7.5) }, sourceSlideSize: { width: emu(13.333), height: emu(7.5) }, designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] }, slides: [0, 1, 2].map((index) => ({ id: `slide-${index + 1}`, slideNumber: index + 1, sourceSlideId: `source-${index + 1}`, sourceTextHash: "c".repeat(64), contentCoverage: { exactTextMapped: true, sourceCharacterCount: 1, mappedCharacterCount: 1, sourceTextBoxCount: 1, mappedTextNodeCount: 1, groupedOrUnsupportedTextPresent: false }, sourceRevision: "source", recipe: "ornl-title-content", background: "#FFFFFF", status: "designed", designRationale: "test", figureTreatments: [], nodes: [titles[index], bodies[index], ...(index < 2 ? [tables[index]] : [])], updatedAt: "2026-08-17T00:00:00.000Z" })) };
  const review = analyzeStudioDeckConsistency(scene);
  assert.equal(review.issueCount, 3);
  assert.deepEqual(review.issues.map((issue) => issue.category).sort(), ["component-type", "table-system", "title-grid"]);
  assert.deepEqual(review.issues.find((issue) => issue.category === "title-grid")?.slideNumbers, [3]);
  assert.deepEqual(review.issues.find((issue) => issue.category === "component-type")?.slideNumbers, [3]);
  assert.deepEqual(review.issues.find((issue) => issue.category === "table-system")?.slideNumbers, [2]);
});

test("approved template-layout covers do not become false title-grid outliers", () => {
  const coverTitle = node("cover-title", "title", 1.2);
  const contentTitles = [node("title-1", "title"), node("title-2", "title"), node("title-3", "title")];
  const slides = [coverTitle, ...contentTitles].map((title, index) => ({ id: `slide-${index + 1}`, slideNumber: index + 1, sourceSlideId: `source-${index + 1}`, sourceTextHash: "c".repeat(64), contentCoverage: { exactTextMapped: true, sourceCharacterCount: 1, mappedCharacterCount: 1, sourceTextBoxCount: 1, mappedTextNodeCount: 1, groupedOrUnsupportedTextPresent: false }, sourceRevision: "source", recipe: index === 0 ? "template-layout" as const : "ornl-title-content" as const, targetLayoutId: index === 0 ? "layout-1" : undefined, background: "#FFFFFF", status: "designed" as const, designRationale: "test", figureTreatments: [], nodes: [title], updatedAt: "2026-08-17T00:00:00.000Z" }));
  const scene: StudioWebScene = { schema: STUDIO_WEB_SCENE_SCHEMA, version: STUDIO_WEB_SCENE_VERSION, revision: "cover-grid", deckId: "deck", sourceSha256: "b".repeat(64), slideSize: { width: emu(13.333), height: emu(7.5) }, sourceSlideSize: { width: emu(13.333), height: emu(7.5) }, designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] }, slides };
  assert.equal(analyzeStudioDeckConsistency(scene).issues.some((issue) => issue.category === "title-grid"), false);
});
