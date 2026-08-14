import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const capture = path.join(os.tmpdir(), `presentation-studio-smoke-${process.pid}.png`);
const electronBin = path.join(root, "node_modules", "electron", "cli.js");
const child = spawn(process.execPath, [electronBin, "."], {
  cwd: root,
  env: { ...process.env, PRESENTATION_STUDIO_SMOKE_TEST: "1", PRESENTATION_STUDIO_CAPTURE_PATH: capture },
  // The smoke flow drives the renderer itself. Do not let terminal input
  // accidentally advance the focused onboarding control while it is measured.
  stdio: ["ignore", "inherit", "inherit"],
});
const code = await new Promise((resolve) => child.on("exit", resolve));
if (code !== 0) process.exit(code ?? 1);
const stats = await fs.stat(capture);
const png = await fs.readFile(capture);
if (png.length < 24) throw new Error("Electron smoke capture is truncated.");
const signature = png.subarray(0, 8).toString("hex");
const width = png.readUInt32BE(16);
const height = png.readUInt32BE(20);
if (signature !== "89504e470d0a1a0a") throw new Error("Electron smoke capture is not a valid PNG image.");
if (width < 1200 || height < 700) throw new Error(`Electron smoke capture is unexpectedly small (${width} x ${height}).`);
// The flat ORNL UI compresses efficiently, so compressed byte size is only a
// coarse blank-image guard. Dimension checks above remain stable as styling
// changes make otherwise healthy screenshots larger or smaller on disk.
if (stats.size < 20_000) throw new Error("Electron smoke capture may be blank or incomplete.");
console.log(`Electron smoke passed with a ${width} x ${height}, ${stats.size.toLocaleString()}-byte renderer capture.`);
