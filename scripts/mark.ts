import fs from "node:fs";
import path from "node:path";

const note = process.argv.slice(2).join(" ").trim();
if (!note) {
  console.error("usage: npm run mark -- \"<what you just did in Teams>\"");
  process.exit(1);
}

const dir = path.resolve("captures");
if (!fs.existsSync(dir)) {
  console.error("no captures/ directory — start a capture first");
  process.exit(1);
}
const files = fs
  .readdirSync(dir)
  .filter((f) => f.endsWith(".jsonl"))
  .map((f) => ({ f, m: fs.statSync(path.join(dir, f)).mtimeMs }))
  .sort((a, b) => b.m - a.m);
if (files.length === 0) {
  console.error("no capture file found");
  process.exit(1);
}
const target = path.join(dir, files[0].f);
fs.appendFileSync(
  target,
  JSON.stringify({ kind: "mark", ts: new Date().toISOString(), note }) + "\n",
);
console.error(`marked ${target}: ${note}`);
