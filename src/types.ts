export interface Cookie {
  name: string;
  value: string;
  domain: string;
  path?: string;
  secure?: boolean;
  httpOnly?: boolean;
  expires?: number;
}

export interface SessionExport {
  cookies: Cookie[];
  /** Bearer tokens captured from the Teams app at login, keyed by JWT audience */
  tokens?: Record<string, string>;
  exportedAt: string;
  source?: string;
}

export type ParamLocation = "path" | "query" | "header" | "body";

export interface EndpointParam {
  name: string;
  in: ParamLocation;
  required?: boolean;
  description?: string;
}

export type AuthMode = "cookie" | "bearer" | "cookie+bearer" | "none";
export type BodyType = "json" | "form" | "multipart" | "binary" | "none";
export type HttpMethod = "GET" | "POST" | "PUT" | "PATCH" | "DELETE" | "HEAD";

export interface EndpointSpec {
  method: HttpMethod;
  url: string;
  params?: EndpointParam[];
  bodyType?: BodyType;
  auth?: AuthMode;
  tokenSource?: string;
  headers?: Record<string, string>;
  notes?: string;
  source?: string;
}

export interface EndpointRegistry {
  version: number;
  updated: string;
  endpoints: Record<string, EndpointSpec>;
}

export interface ApiResponse {
  status: number;
  headers: Headers;
  json?: unknown;
  text?: string;
  buffer?: Buffer;
}
