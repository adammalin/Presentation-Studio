import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";

const require = createRequire(import.meta.url);
const { buildEditorContextMenuTemplate, uniqueSuggestions } = require("../electron/editor-context-menu.cjs") as {
  buildEditorContextMenuTemplate: (params: Record<string, unknown>, actions: { replaceMisspelling: (word: string) => void; addToDictionary: (word: string) => void }) => Array<Record<string, unknown>>;
  uniqueSuggestions: (values: unknown[]) => string[];
};

test("editable spelling context menu exposes bounded replacement suggestions and dictionary action", () => {
  const replacements: string[] = [];
  const dictionary: string[] = [];
  const template = buildEditorContextMenuTemplate({
    isEditable: true,
    misspelledWord: "alignement",
    dictionarySuggestions: ["alignment", "alinement", "alignment", ""],
    editFlags: { canUndo: true, canRedo: false, canCut: true, canCopy: true, canPaste: true, canSelectAll: true },
  }, {
    replaceMisspelling: (word) => replacements.push(word),
    addToDictionary: (word) => dictionary.push(word),
  });
  assert.deepEqual(template.slice(0, 2).map((item) => item.label), ["alignment", "alinement"]);
  (template[0].click as () => void)();
  (template.find((item) => String(item.label).startsWith("Add “"))?.click as () => void)();
  assert.deepEqual(replacements, ["alignment"]);
  assert.deepEqual(dictionary, ["alignement"]);
  assert.equal(template.some((item) => item.role === "paste"), true);
  assert.equal(template.find((item) => item.role === "redo")?.enabled, false);
});

test("context menu stays editing-only and limits duplicate suggestions", () => {
  assert.deepEqual(buildEditorContextMenuTemplate({ isEditable: false }, { replaceMisspelling() {}, addToDictionary() {} }), []);
  assert.deepEqual(uniqueSuggestions(["one", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]), ["one", "two", "three", "four", "five", "six", "seven", "eight"]);
});
