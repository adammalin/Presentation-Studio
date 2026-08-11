import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const serverPath = path.join(root, "mcp", "server.mjs");
const entry = { command: process.execPath, args: [serverPath] };
const snippet = { mcpServers: { "presentation-studio": entry } };
const writeIndex = process.argv.indexOf("--write");

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
