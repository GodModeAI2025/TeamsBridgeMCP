import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");

const DEFAULT_DOMAINS = [
  "https://teams.microsoft.com",
  "https://skypeforbusiness.online",
  "https://web.skype.com",
  "https://skype.com",
  "https://login.microsoftonline.com",
  "https://account.live.com",
  "https://office.com",
];

async function main() {
  const argv = process.argv.slice(2);
  const get = (flag: string, dflt?: string): string | undefined => {
    const i = argv.indexOf(flag);
    return i >= 0 ? argv[i + 1] : dflt;
  };
  const port = get("--port", "9222");
  const out = path.resolve(get("--out", path.join(root, "data", "session.json")));
  const domainsArg = get("--domains");
  const urls = (domainsArg
    ? domainsArg.split(",").map((d) => d.trim())
    : DEFAULT_DOMAINS
  ).map((d) => (d.startsWith("http") ? d : `https://${d}`));

  const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`);
  const ctx = browser.contexts()[0];
  if (!ctx) throw new Error("no browser context found via CDP");

  const cookies = await ctx.cookies(...urls);
  const exportData = {
    cookies,
    exportedAt: new Date().toISOString(),
    source: `cdp:127.0.0.1:${port}`,
  };
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(exportData, null, 2));
  fs.chmodSync(out, 0o600);

  const names = [...new Set(cookies.map((c) => c.name))].sort();
  console.log(`exported ${cookies.length} cookies -> ${out}`);
  console.log(`cookie names: ${names.join(", ") || "(none)"}`);
  const hasToken = cookies.some((c) => /access_token|token/i.test(c.name));
  if (!hasToken) {
    console.warn("warning: no token-like cookie found — the Teams session may rely on in-memory tokens; capture traffic to verify auth flow");
  }
  try {
    await browser.close();
  } catch {
    // ignore
  }
}

main().catch((e) => {
  console.error(String(e?.stack ?? e));
  process.exit(1);
});
