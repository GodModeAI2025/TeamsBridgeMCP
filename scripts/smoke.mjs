import { spawn } from "node:child_process";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { expandUrl, setDetected } from "../dist/registry.js";
import { htmlToText, toTeamsHtml } from "../dist/format.js";
import { detectTenantValues, staleBrowserPids } from "../dist/login.js";
import { buildUrl } from "../dist/http.js";

process.env.TEAMS_ONEDRIVE_HOST = "contoso-my.sharepoint.com";
assert.equal(expandUrl("https://${TEAMS_ONEDRIVE_HOST}/x?$select=id"), "https://contoso-my.sharepoint.com/x?$select=id");
assert.equal(expandUrl("https://h/${TEAMS_REGION}/{conversationId}"), "https://h/emea/{conversationId}");
delete process.env.TEAMS_ONEDRIVE_HOST;
assert.throws(() => expandUrl("https://${TEAMS_ONEDRIVE_HOST}/"), /TEAMS_ONEDRIVE_HOST/);

// tenant values from login traffic: flag/env > detected > default
setDetected({ TEAMS_REGION: "amer" });
assert.equal(expandUrl("https://h/${TEAMS_REGION}"), "https://h/amer");
process.env.TEAMS_REGION = "apac";
assert.equal(expandUrl("https://h/${TEAMS_REGION}"), "https://h/apac");
delete process.env.TEAMS_REGION;
setDetected({});
assert.deepEqual(detectTenantValues("https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME"), { TEAMS_REGION: "amer" });
assert.deepEqual(detectTenantValues("https://teams.cloud.microsoft/api/mt/part/amer-03/beta/me"), { TEAMS_PARTITION: "amer-03" });
assert.deepEqual(detectTenantValues("https://contoso-my.sharepoint.com/_api/x", "00000003-0000-0ff1-ce00-000000000000"), { TEAMS_ONEDRIVE_HOST: "contoso-my.sharepoint.com" });
assert.deepEqual(detectTenantValues("https://contoso.sharepoint.com/_api/x", "00000003-0000-0ff1-ce00-000000000000"), {});

// ids made of dots would be resolved as path segments
assert.throws(() => buildUrl("https://h/a/{id}/b", { id: ".." }), /invalid id/);
assert.equal(buildUrl("https://h/a/{id}", { id: "19:x@thread.v2" }), "https://h/a/19%3Ax%40thread.v2");

// markdown subset -> Teams HTML, escaped, links only http(s)/mailto
assert.equal(toTeamsHtml("**fett** und *kursiv* und ~~weg~~").content, "<p><b>fett</b> und <i>kursiv</i> und <s>weg</s></p>");
assert.equal(toTeamsHtml("a <b> & `x_y_z`").content, "<p>a &lt;b&gt; &amp; <code>x_y_z</code></p>");
assert.equal(toTeamsHtml("[docs](https://a.com/x_y_z) [bad](javascript:alert(1))").content, '<p><a href="https://a.com/x_y_z">docs</a> [bad](javascript:alert(1))</p>');
assert.equal(toTeamsHtml("Liste:\n- eins\n- zwei\n\n30. Oktober\n2. b").content, "<p>Liste:</p><ul><li>eins</li><li>zwei</li></ul><p>30. Oktober<br>2. b</p>");
assert.equal(toTeamsHtml("see https://a.com/x_y*z*?a=1&b=2.").content, '<p>see <a href="https://a.com/x_y*z*?a=1&amp;b=2">https://a.com/x_y*z*?a=1&amp;b=2</a>.</p>');
assert.equal(toTeamsHtml("<https://a.com>").content, '<p>&lt;<a href="https://a.com">https://a.com</a>&gt;</p>');
assert.equal(toTeamsHtml("zeile1\nzeile2").content, "<p>zeile1<br>zeile2</p>");
const m = toTeamsHtml("hi @alex_b@contoso.com", new Map([["alex_b@contoso.com", { objectId: "1111", displayName: "Alex <B>" }]]));
assert.equal(m.content, '<p>hi <span itemscope itemtype="http://schema.skype.com/Mention" itemid="0">Alex &lt;B&gt;</span></p>');
assert.equal(m.mentions[0].mri, "8:orgid:1111");
assert.equal(htmlToText("<p>a &amp; b</p><p>c<br>d</p><ul><li>x</li></ul>"), "a & b\nc\nd\n- x");

// a leftover browser on our profile is found by its --user-data-dir (ps on macOS/Linux, CIM on Windows)
const fakeDir = path.join(os.tmpdir(), "teams-smoke-profile");
const fake = spawn(process.execPath, ["-e", "setTimeout(() => {}, 30000)", "--", `--user-data-dir=${fakeDir}`]);
await new Promise((r) => setTimeout(r, 300));
assert.ok(staleBrowserPids(fakeDir).includes(fake.pid));
fake.kill();

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
