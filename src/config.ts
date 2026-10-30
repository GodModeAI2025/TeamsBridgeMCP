import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.resolve(here, "..");

export function sessionFile(): string {
  const env = process.env.TEAMS_SESSION_FILE;
  if (env) return path.resolve(env);
  return path.join(ROOT, "data", "session.json");
}

export function registryFile(): string {
  const env = process.env.TEAMS_ENDPOINTS_FILE;
  if (env) return path.resolve(env);
  return path.join(ROOT, "data", "endpoints.json");
}
