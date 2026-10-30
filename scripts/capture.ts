import fs from "node:fs";
import path from "node:path";
import { chromium, Page } from "playwright-core";

async function main() {
  const argv = process.argv.slice(2);
  const get = (flag: string, dflt?: string): string | undefined => {
    const i = argv.indexOf(flag);
    return i >= 0 ? argv[i + 1] : dflt;
  };
  const port = get("--port", "9222");
  const label = get("--label", "capture");
  const origin = get("--origin");
  const maxBody = parseInt(get("--max-body", "262144"), 10);
  const cdpUrl = `http://127.0.0.1:${port}`;

  const browser = await chromium.connectOverCDP(cdpUrl);
  const ctx = browser.contexts()[0];
  if (!ctx) throw new Error("no browser context found via CDP");

  const outDir = path.resolve("captures");
  fs.mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, `${label}-${new Date().toISOString().replace(/[:.]/g, "-")}.jsonl`);
  fs.writeFileSync(outFile, "");
  fs.chmodSync(outFile, 0o600);

  const t0 = Date.now();
  const matches = (url: string) => !origin || url.includes(origin);

  const clip = (s: string | null | undefined): string | undefined => {
    if (s === null || s === undefined) return undefined;
    return s.length > maxBody ? `${s.slice(0, maxBody)}…[truncated total=${s.length}]` : s;
  };

  const attachPage = (page: Page) => {
    page.on("response", async (res) => {
      try {
        const req = res.request();
        const url = req.url();
        if (!matches(url)) return;
        const resHeaders = res.headers();
        const ctype = resHeaders["content-type"] ?? "";
        let resBody: string | undefined;
        if (/json|text\/|xml/.test(ctype)) {
          try {
            resBody = clip((await res.body()).toString("utf8"));
          } catch {
            resBody = "<unreadable>";
          }
        }
        fs.appendFileSync(
          outFile,
          JSON.stringify({
            kind: "res",
            ts: new Date().toISOString(),
            dt: ((Date.now() - t0) / 1000).toFixed(3),
            page: page.url(),
            method: req.method(),
            url,
            reqHeaders: req.headers(),
            reqBody: clip(req.postData()),
            status: res.status(),
            resContentType: ctype,
            resBody,
          }) + "\n",
        );
      } catch {
        // keep capturing
      }
    });
    page.on("requestfailed", (req) => {
      try {
        const url = req.url();
        if (!matches(url)) return;
        fs.appendFileSync(
          outFile,
          JSON.stringify({
            kind: "fail",
            ts: new Date().toISOString(),
            dt: ((Date.now() - t0) / 1000).toFixed(3),
            page: page.url(),
            method: req.method(),
            url,
            error: req.failure()?.errorText ?? "unknown",
          }) + "\n",
        );
      } catch {
        // keep capturing
      }
    });
  };

  for (const p of ctx.pages()) attachPage(p);
  ctx.on("page", (p) => attachPage(p));

  console.error(`capture started -> ${outFile}`);
  console.error(`origin filter: ${origin ?? "(none)"}`);
  console.error("perform the target actions in the browser; stop with Ctrl+C.");

  const shutdown = async () => {
    try {
      await browser.close();
    } catch {
      // ignore
    }
    process.exit(0);
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
  await new Promise(() => {});
}

main().catch((e) => {
  console.error(String(e?.stack ?? e));
  process.exit(1);
});
