import fs from "node:fs";
import { Cookie, SessionExport } from "./types.js";

export class SessionError extends Error {}

export interface SessionState {
  cookies: Cookie[];
  cookieHeaderFor(origin: string): string;
  token(name: string): string | undefined;
  cookieNames(): string[];
  tokenAudiences(): string[];
  detected: Record<string, string>;
}

export function loadSession(file: string): SessionState {
  if (!fs.existsSync(file)) {
    throw new SessionError(
      `no session export at ${file}. Call the teams_login tool or run: npm run login`,
    );
  }
  let raw: SessionExport;
  try {
    raw = JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (e) {
    throw new SessionError(`session file is not valid JSON: ${(e as Error).message}`);
  }
  const cookies: Cookie[] = Array.isArray(raw.cookies) ? raw.cookies : [];
  const tokens = raw.tokens ?? {};
  if (cookies.length === 0 && Object.keys(tokens).length === 0) {
    throw new SessionError("session export contains no cookies and no tokens");
  }

  return {
    cookies,
    detected: raw.detected ?? {},
    cookieHeaderFor(origin: string): string {
      const host = new URL(origin).hostname;
      return cookies
        .filter((c) => {
          const d = c.domain.replace(/^\./, "");
          return host === d || host.endsWith("." + d);
        })
        .map((c) => `${c.name}=${c.value}`)
        .join("; ");
    },
    token(name: string): string | undefined {
      return tokens[name] ?? cookies.find((c) => c.name === name)?.value;
    },
    tokenAudiences(): string[] {
      return Object.keys(tokens).sort();
    },
    cookieNames(): string[] {
      return [...new Set(cookies.map((c) => c.name))].sort();
    },
  };
}
