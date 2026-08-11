import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const candidates = execFileSync("git", ["ls-files", "--cached", "--others", "--exclude-standard"], { cwd: root, encoding: "utf8" }).split(/\r?\n/).filter(Boolean);
const history = execFileSync("git", ["rev-list", "--objects", "--all"], { cwd: root, encoding: "utf8" }).split(/\r?\n/).map((line) => line.slice(line.indexOf(" ") + 1)).filter(Boolean);
const forbiddenExtensions = /\.(?:pstudio|pstudio-secure|pptx|pptm|potx)$/i;
const forbiddenNames = /(?:mcp-runtime|autosave|recovery|client[-_ ]deck|manuscript)/i;
const exceptions = new Set([]);
const violations = [];

for (const relative of candidates) {
  if (exceptions.has(relative)) continue;
  if (forbiddenExtensions.test(relative) || forbiddenNames.test(relative)) violations.push(relative);
  const absolute = path.join(root, relative);
  const stats = fs.statSync(absolute);
  if (stats.size > 30 * 1024 * 1024) violations.push(`${relative} (over 30 MB)`);
}

for (const relative of history) {
  if (forbiddenExtensions.test(relative) || forbiddenNames.test(relative)) violations.push(`${relative} (Git history)`);
}

if (violations.length) {
  console.error("Repository data-safety scan rejected tracked working/client artifacts:");
  for (const violation of [...new Set(violations)]) console.error(`- ${violation}`);
  process.exit(1);
}
console.log(`Repository data-safety scan passed for ${candidates.length} tracked or candidate files and ${history.length} historical paths.`);
