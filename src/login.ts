import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { chromium } from "playwright-core";
import { ROOT } from "./config.js";

export const COOKIE_DOMAINS = [
  "https://teams.microsoft.com",
  "https://teams.cloud.microsoft",
  "https://skypeforbusiness.online",
  "https://web.skype.com",
  "https://skype.com",
  "https://login.microsoftonline.com",
  "https://account.live.com",
  "https://office.com",
];

const SHAREPOINT_AUDIENCE = "00000003-0000-0ff1-ce00-000000000000";

// Token audiences the registry endpoints use as tokenSource: chats/messages (chatsvc), profiles, files (SharePoint)
export const REQUIRED_AUDIENCES = [
  "https://ic3.teams.office.com",
  "https://chatsvcagg.teams.microsoft.com",
  "https://api.spaces.skype.com",
  SHAREPOINT_AUDIENCE,
];

// Substrate search token: the app only requests it once the user searches
const SEARCH_AUDIENCE = "66a88757-258c-4c72-893c-3e8bed4d6899";

// Persistent profile: after the first login, Microsoft SSO usually logs in again without prompting.
const PROFILE_DIR = path.join(ROOT, "data", "browser-profile");

/**
 * PIDs of browsers still running on our profile, e.g. after the host crashed during a sign-in.
 * Such a process holds the profile lock, and every later launch fails until it is gone.
 */
export function staleBrowserPids(profileDir = PROFILE_DIR): number[] {
  const win = process.platform === "win32";
  // single quotes only: powershell.exe -Command tends to strip embedded double quotes
  const out = win
    ? execFileSync("powershell.exe", ["-NoProfile", "-Command", "Get-CimInstance Win32_Process | ForEach-Object { '{0} {1}' -f $_.ProcessId, $_.CommandLine }"], { encoding: "utf8" })
    : execFileSync("ps", ["-axww", "-o", "pid=,command="], { encoding: "utf8" });
  const needle = `--user-data-dir=${profileDir}`;
  return out
    .split(/\r?\n/)
    .filter((l) => (win ? l.toLowerCase().includes(needle.toLowerCase()) : l.includes(needle)))
    .map((l) => Number(l.trim().split(/\s+/)[0]))
    .filter((pid) => pid > 0 && pid !== process.pid);
}

async function launch(headless: boolean) {
  const opts = {
    channel: process.env.TEAMS_BROWSER || "chrome",
    headless,
    viewport: null,
    // Edge on macOS otherwise writes its HTTP cache to ~/Library/Caches, next to the user's own Edge, where deleting data/ misses it
    args: [`--disk-cache-dir=${path.join(PROFILE_DIR, "cache")}`],
  };
  try {
    return await chromium.launchPersistentContext(PROFILE_DIR, opts);
  } catch (e) {
    let pids: number[] = [];
    try {
      pids = staleBrowserPids();
    } catch {
      // process listing failed: report the launch error, not this one
    }
    if (!pids.length) throw e;
    for (const pid of pids) {
      try {
        process.kill(pid);
      } catch {
        // already gone
      }
    }
    await new Promise((r) => setTimeout(r, 2000));
    return chromium.launchPersistentContext(PROFILE_DIR, opts);
  }
}

/** Returns the JWT `aud` claim of a "Bearer <jwt>" header value, or undefined. */
export function bearerAudience(header: string | undefined): string | undefined {
  const m = header?.match(/^Bearer ([\w-]+)\.([\w-]+)\.[\w-]*$/);
  if (!m) return undefined;
  try {
    const aud = JSON.parse(Buffer.from(m[2], "base64url").toString("utf8")).aud;
    return typeof aud === "string" ? aud : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Tenant values the registry needs as ${NAME}, read from the Teams client's own requests:
 * chat service region, middle-tier partition, and the OneDrive host (the request that carries the SharePoint token).
 */
export function detectTenantValues(url: string, aud?: string): Record<string, string> {
  const out: Record<string, string> = {};
  const region = url.match(/\/api\/chatsvc\/([a-z]+)\/v1\//)?.[1];
  if (region) out.TEAMS_REGION = region;
  const partition = url.match(/\/api\/mt\/part\/([a-z0-9-]+)\//)?.[1];
  if (partition) out.TEAMS_PARTITION = partition;
  if (aud === SHAREPOINT_AUDIENCE) {
    const host = new URL(url).hostname;
    if (host.endsWith("-my.sharepoint.com")) out.TEAMS_ONEDRIVE_HOST = host;
  }
  return out;
}

/**
 * Opens Chrome on Teams, waits until the app has sent requests carrying the required bearer tokens,
 * then writes cookies + tokens to `out`.
 */
export async function browserLogin(
  out: string,
  { headless = false, timeoutMs = 5 * 60_000 } = {},
): Promise<{ tokenAudiences: string[]; out: string }> {
  const ctx = await launch(headless);
  const tokens: Record<string, string> = {};
  const detected: Record<string, string> = {};
  ctx.on("request", (req) => {
    const auth = req.headers()["authorization"];
    const aud = bearerAudience(auth);
    if (aud) tokens[aud] = auth.slice("Bearer ".length);
    Object.assign(detected, { ...detectTenantValues(req.url(), aud), ...detected });
  });
  try {
    const page = ctx.pages()[0] ?? (await ctx.newPage());
    await page.goto("https://teams.microsoft.com", { waitUntil: "domcontentloaded" });

    const deadline = Date.now() + timeoutMs;
    while (!REQUIRED_AUDIENCES.every((a) => tokens[a])) {
      if (page.isClosed()) throw new Error("browser window was closed before login finished");
      if (Date.now() > deadline) {
        const missing = REQUIRED_AUDIENCES.filter((a) => !tokens[a]).join(", ");
        throw new Error(`login not completed within ${Math.round(timeoutMs / 1000)} s (missing tokens: ${missing})`);
      }
      await new Promise((r) => setTimeout(r, 1000));
    }

    // shortcut: types into the Teams search box to get the search token; breaks if the box selector changes (search then fails, login still works)
    try {
      const box = page.locator("input[type=search], #ms-searchux-input").first();
      await box.click({ timeout: 15_000 });
      await box.fill("teams");
      await box.press("Enter");
      for (let i = 0; i < 10 && !tokens[SEARCH_AUDIENCE]; i++) await page.waitForTimeout(1000);
    } catch {
      // search token is optional
    }

    const cookies = await ctx.cookies(...COOKIE_DOMAINS);
    fs.mkdirSync(path.dirname(out), { recursive: true });
    fs.writeFileSync(
      out,
      JSON.stringify({ cookies, tokens, detected, exportedAt: new Date().toISOString(), source: "playwright-login" }, null, 2),
    );
    fs.chmodSync(out, 0o600);
    return { tokenAudiences: Object.keys(tokens).sort(), out };
  } finally {
    await ctx.close().catch(() => {});
  }
}
