import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("automatic project launch accepts Resources-only projects and enables the current AI access control", async () => {
  const source = await readFile(path.join(root, "electron", "main.cjs"), "utf8");
  assert.match(source, /all embedded resource hashes passed validation\./);
  assert.doesNotMatch(source, /1 decks · preserve-exact/);
  assert.match(source, /textContent\?\.includes\('AI access'\)/);
});
