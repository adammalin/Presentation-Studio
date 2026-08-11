import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const capture = path.join(os.tmpdir(), `presentation-studio-smoke-${process.pid}.png`);
const electronBin = path.join(root, "node_modules", "electron", "cli.js");
const child = spawn(process.execPath, [electronBin, "."], { cwd: root, env: { ...process.env, PRESENTATION_STUDIO_SMOKE_TEST: "1", PRESENTATION_STUDIO_CAPTURE_PATH: capture }, stdio: "inherit" });
const code = await new Promise((resolve) => child.on("exit", resolve));
if (code !== 0) process.exit(code ?? 1);
const stats = await fs.stat(capture);
if (stats.size < 75_000) throw new Error("Electron smoke capture was unexpectedly small and may be blank.");
console.log(`Electron smoke passed with a ${stats.size.toLocaleString()}-byte renderer capture.`);
