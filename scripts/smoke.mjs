import { spawn } from "node:child_process";
import assert from "node:assert/strict";
import { expandUrl } from "../dist/registry.js";

process.env.TEAMS_ONEDRIVE_HOST = "contoso-my.sharepoint.com";
assert.equal(expandUrl("https://${TEAMS_ONEDRIVE_HOST}/x?$select=id"), "https://contoso-my.sharepoint.com/x?$select=id");
assert.equal(expandUrl("https://h/${TEAMS_REGION}/{conversationId}"), "https://h/emea/{conversationId}");
delete process.env.TEAMS_ONEDRIVE_HOST;
assert.throws(() => expandUrl("https://${TEAMS_ONEDRIVE_HOST}/"), /TEAMS_ONEDRIVE_HOST/);

const child = process.argv[2] ? spawn(process.execPath, [process.argv[2]]) : spawn("node", ["dist/index.js"]);
let buf = "";
const pending = new Map();

child.stdout.on("data", (d) => {
  buf += d.toString();
  let idx;
  while ((idx = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, idx).trim();
    buf = buf.slice(idx + 1);
    if (!line) continue;
    const msg = JSON.parse(line);
    if (msg.id !== undefined && pending.has(msg.id)) {
      pending.get(msg.id)(msg);
      pending.delete(msg.id);
    }
  }
});
child.stderr.on("data", (d) => process.stderr.write(`[server] ${d}`));

let nextId = 1;
function request(method, params) {
  const id = nextId++;
  return new Promise((resolve, reject) => {
    pending.set(id, resolve);
    child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n");
    setTimeout(() => reject(new Error(`timeout waiting for ${method}`)), 8000);
  });
}
function notify(method, params) {
  child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method, params }) + "\n");
}

const init = await request("initialize", {
  protocolVersion: "2025-03-26",
  capabilities: {},
  clientInfo: { name: "smoke-test", version: "0.0.1" },
});
console.log("initialize:", JSON.stringify(init.result?.serverInfo));
notify("notifications/initialized", {});

const tools = await request("tools/list", {});
console.log("tools:", tools.result.tools.map((t) => t.name).join(", "));

const status = await request("tools/call", { name: "teams_status", arguments: {} });
console.log("teams_status isError:", status.result?.isError ?? false);
console.log(status.result?.content?.[0]?.text ?? JSON.stringify(status.result));

child.kill();
process.exit(0);
