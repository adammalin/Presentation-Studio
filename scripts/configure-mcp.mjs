import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const serverPath = path.join(root, "mcp", "server.mjs");
const entry = { command: process.execPath, args: [serverPath] };
const snippet = { mcpServers: { "presentation-studio": entry } };
const writeIndex = process.argv.indexOf("--write");
const codexIndex = process.argv.indexOf("--codex");
const CODEX_START = "# BEGIN PRESENTATION STUDIO MCP - managed by installer";
const CODEX_END = "# END PRESENTATION STUDIO MCP - managed by installer";

function configureCodex(target) {
  let current = fs.existsSync(target) ? fs.readFileSync(target, "utf8") : "";
  const startIndex = current.indexOf(CODEX_START);
  const endIndex = current.indexOf(CODEX_END);
  if ((startIndex === -1) !== (endIndex === -1) || endIndex !== -1 && endIndex < startIndex) {
    throw new Error(`Cannot safely update the incomplete Presentation Studio block in ${target}.`);
  }
  if (startIndex !== -1) {
    current = `${current.slice(0, startIndex).trimEnd()}\n${current.slice(endIndex + CODEX_END.length).trimStart()}`;
  }
  const block = [
    CODEX_START,
    "[mcp_servers.presentation_studio]",
    `command = ${JSON.stringify(entry.command)}`,
    `args = [${entry.args.map((item) => JSON.stringify(item)).join(", ")}]`,
    CODEX_END,
  ].join("\n");
  const next = `${current.trimEnd()}${current.trim() ? "\n\n" : ""}${block}\n`;
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, next, { mode: 0o600 });
  console.log(`Configured Presentation Studio in Codex: ${target}`);
}

if (codexIndex !== -1) {
  const supplied = process.argv[codexIndex + 1];
  const target = supplied && !supplied.startsWith("--") ? path.resolve(supplied) : path.join(os.homedir(), ".codex", "config.toml");
  configureCodex(target);
  process.exit(0);
}

if (writeIndex === -1) {
  console.log(JSON.stringify(snippet, null, 2));
  console.log("\nAdd the presentation-studio entry to any MCP client's standard mcpServers configuration. Presentation Studio remains model-independent.");
  process.exit(0);
}

const targetArg = process.argv[writeIndex + 1];
if (!targetArg) throw new Error("Pass an explicit MCP JSON configuration path after --write.");
const target = path.resolve(targetArg);
let current = {};
if (fs.existsSync(target)) current = JSON.parse(fs.readFileSync(target, "utf8"));
const next = { ...current, mcpServers: { ...(current.mcpServers ?? {}), "presentation-studio": entry } };
fs.mkdirSync(path.dirname(target), { recursive: true });
fs.writeFileSync(target, `${JSON.stringify(next, null, 2)}\n`, { mode: 0o600 });
console.log(`Configured Presentation Studio in ${target}`);
