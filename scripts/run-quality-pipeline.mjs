import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";
const flags = new Set(process.argv.slice(2));
const includeDesktopSmoke = flags.has("--desktop-smoke") && !flags.has("--ci");
const includeNativeCanary = flags.has("--native-canary") && !flags.has("--ci");
const reportPath = path.join(root, "tmp", "quality-pipeline", "latest.json");

const steps = [
  { id: "lint", args: ["run", "lint"] },
  { id: "test", args: ["test"] },
  { id: "build", args: ["run", "build"] },
  { id: "data-safety", args: ["run", "check:data-safety"] },
  ...(includeDesktopSmoke ? [{ id: "desktop-smoke", args: ["run", "desktop:smoke"] }] : []),
  ...(includeNativeCanary ? [{ id: "native-canary", args: ["run", "test:canary:native"] }] : []),
];

fs.mkdirSync(path.dirname(reportPath), { recursive: true, mode: 0o700 });
const startedAt = new Date().toISOString();
const results = [];
let failed = false;
for (const step of steps) {
  const started = Date.now();
  process.stdout.write(`\n[quality] ${step.id}\n`);
  const result = spawnSync(npmCommand, step.args, { cwd: root, stdio: "inherit", env: process.env });
  const exitCode = result.status ?? 1;
  results.push({ id: step.id, command: `npm ${step.args.join(" ")}`, exitCode, durationMs: Date.now() - started });
  if (exitCode !== 0) {
    failed = true;
    break;
  }
}

const report = {
  schema: "presentation-studio/quality-pipeline",
  version: 1,
  startedAt,
  completedAt: new Date().toISOString(),
  status: failed ? "failed" : "passed",
  mode: includeNativeCanary ? "powerpoint-native" : includeDesktopSmoke ? "desktop" : flags.has("--ci") ? "portable-ci" : "local",
  clientContentIncluded: false,
  note: includeNativeCanary
    ? "The native lane uses only the synthetic precision-layout canary. Client presentations and qualification images are never uploaded or archived by this script."
    : "Portable checks do not establish PowerPoint-native visual quality. Run the native lane on an approved self-hosted workstation before release.",
  results,
};
fs.writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
process.stdout.write(`\n[quality] ${report.status} · ${reportPath}\n`);
if (failed) process.exitCode = 1;
