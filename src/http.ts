import { ApiResponse, EndpointSpec } from "./types.js";
import { SessionState } from "./session.js";

export class ApiError extends Error {
  constructor(
    public status: number,
    message: string,
    public body?: string,
  ) {
    super(message);
  }
}

export function buildUrl(base: string, params?: Record<string, unknown>): string {
  let s = base;
  const rest: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(params ?? {})) {
    if (v === undefined || v === null) continue;
    const ph = `{${k}}`;
    if (s.includes(ph)) s = s.replace(ph, encodeURIComponent(String(v)));
    else rest[k] = v;
  }
  const u = new URL(s);
  for (const [k, v] of Object.entries(rest)) {
    if (v !== undefined && v !== null) u.searchParams.set(k, String(v));
  }
  return u.toString();
}

interface RequestOptions {
  params?: Record<string, unknown>;
  body?: unknown;
  timeoutMs?: number;
  raw?: boolean;
}

export async function apiRequest(
  spec: EndpointSpec,
  session: SessionState,
  opts: RequestOptions = {},
): Promise<ApiResponse> {
  const url = buildUrl(spec.url, opts.params);
  const authMode = spec.auth ?? "cookie";
  const origin = new URL(spec.url).origin;
  const headers: Record<string, string> = { ...(spec.headers ?? {}) };

  if (authMode === "cookie" || authMode === "cookie+bearer") {
    const ch = session.cookieHeaderFor(origin);
    if (ch) headers["Cookie"] = ch;
  }
  if (authMode === "bearer" || authMode === "cookie+bearer") {
    if (!spec.tokenSource) {
      throw new ApiError(0, `endpoint declares bearer auth but has no tokenSource`);
    }
    const token = session.token(spec.tokenSource);
    if (!token) {
      throw new ApiError(401, `no token '${spec.tokenSource}' in session; call teams_login`);
    }
    headers["Authorization"] = `Bearer ${token}`;
  }

  const init: RequestInit = { method: spec.method, headers, redirect: "follow" };
  if (opts.body !== undefined) {
    const bodyType = spec.bodyType ?? "json";
    if (bodyType === "json") {
      headers["Content-Type"] = headers["Content-Type"] ?? "application/json; charset=utf-8";
      init.body = JSON.stringify(opts.body);
    } else if (bodyType === "form") {
      headers["Content-Type"] = headers["Content-Type"] ?? "application/x-www-form-urlencoded";
      const p = new URLSearchParams();
      for (const [k, v] of Object.entries(opts.body as Record<string, unknown>)) p.set(k, String(v));
      init.body = p.toString();
    } else if (bodyType === "binary") {
      headers["Content-Type"] = headers["Content-Type"] ?? "application/octet-stream";
      init.body = new Uint8Array(opts.body as Buffer);
    } else if (bodyType === "multipart") {
      init.body = opts.body as FormData;
    }
  }

  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), opts.timeoutMs ?? 60000);
  try {
    const res = await fetch(url, { ...init, signal: ctrl.signal });
    if (res.status === 401) {
      throw new ApiError(401, "HTTP 401: token rejected or expired", (await res.text()).slice(0, 400));
    }
    // 403 is not a login problem: signing in again does not help, so the caller must not retry with a new token
    if (res.status === 403) {
      throw new ApiError(
        403,
        "HTTP 403: forbidden (no access to this conversation or resource, or wrong region/partition)",
        (await res.text()).slice(0, 400),
      );
    }
    if (!res.ok) {
      const bodyText = (await res.text()).slice(0, 2000);
      throw new ApiError(res.status, `HTTP ${res.status}`, bodyText);
    }

    const respHeaders = res.headers;

    if (opts.raw) {
      const buf = Buffer.from(await res.arrayBuffer());
      return { status: res.status, headers: respHeaders, buffer: buf };
    }
    const text = await res.text();
    let json: unknown;
    try {
      json = JSON.parse(text);
    } catch {
      json = undefined;
    }
    return {
      status: res.status,
      headers: respHeaders,
      json: json === undefined ? undefined : json,
      text: json === undefined ? text : undefined,
    };
  } catch (e) {
    if (e instanceof ApiError) throw e;
    if ((e as Error).name === "AbortError") throw new ApiError(0, "request timed out");
    throw new ApiError(0, `request failed: ${(e as Error).message}`);
  } finally {
    clearTimeout(timer);
  }
}
