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

// Token audiences the registry endpoints use as tokenSource: chats/messages (chatsvc), profiles, files (SharePoint)
export const REQUIRED_AUDIENCES = [
  "https://ic3.teams.office.com",
  "https://chatsvcagg.teams.microsoft.com",
  "https://api.spaces.skype.com",
  "00000003-0000-0ff1-ce00-000000000000",
];

// Substrate search token: the app only requests it once the user searches
const SEARCH_AUDIENCE = "66a88757-258c-4c72-893c-3e8bed4d6899";

// Persistent profile: after the first login, Microsoft SSO usually logs in again without prompting.
const PROFILE_DIR = path.join(ROOT, "data", "browser-profile");

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
 * Opens Chrome on Teams, waits until the app has sent requests carrying the required bearer tokens,
 * then writes cookies + tokens to `out`.
 */
export async function browserLogin(
  out: string,
  { headless = false, timeoutMs = 5 * 60_000 } = {},
): Promise<{ tokenAudiences: string[]; out: string }> {
  const ctx = await chromium.launchPersistentContext(PROFILE_DIR, {
    channel: process.env.TEAMS_BROWSER || "chrome",
    headless,
    viewport: null,
  });
  const tokens: Record<string, string> = {};
  ctx.on("request", (req) => {
    const auth = req.headers()["authorization"];
    const aud = bearerAudience(auth);
    if (aud) tokens[aud] = auth.slice("Bearer ".length);
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
      JSON.stringify({ cookies, tokens, exportedAt: new Date().toISOString(), source: "playwright-login" }, null, 2),
    );
    fs.chmodSync(out, 0o600);
    return { tokenAudiences: Object.keys(tokens).sort(), out };
  } finally {
    await ctx.close().catch(() => {});
  }
}
