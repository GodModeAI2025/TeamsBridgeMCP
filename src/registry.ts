import fs from "node:fs";
import { EndpointRegistry, EndpointSpec } from "./types.js";

export class RegistryError extends Error {}

export function loadRegistry(file: string): EndpointRegistry {
  if (!fs.existsSync(file)) {
    throw new RegistryError(`endpoint registry not found at ${file}`);
  }
  let reg: EndpointRegistry;
  try {
    reg = JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (e) {
    throw new RegistryError(`endpoint registry is not valid JSON: ${(e as Error).message}`);
  }
  if (!reg || typeof reg.endpoints !== "object") {
    throw new RegistryError("endpoint registry missing 'endpoints' object");
  }
  return reg;
}

// Tenant-specific URL parts in the registry as ${NAME}, filled from the environment
const URL_DEFAULTS: Record<string, string> = {
  TEAMS_REGION: "emea",
  TEAMS_PARTITION: "emea-02",
  TEAMS_CHAT_FILES_FOLDER: "Microsoft Teams Chat Files",
};

// Values the login read from the Teams client's traffic; flags and environment still win
let detected: Record<string, string> = {};
export function setDetected(values: Record<string, string>): void {
  detected = values;
}

export function expandUrl(url: string): string {
  return url.replace(/\$\{(\w+)\}/g, (_, name: string) => {
    const v = process.env[name] || detected[name] || URL_DEFAULTS[name];
    if (!v) throw new RegistryError(`environment variable ${name} is not set (see README, "Configuration")`);
    return v;
  });
}

export function endpoint(reg: EndpointRegistry, key: string): EndpointSpec {
  const spec = reg.endpoints[key];
  if (!spec) {
    throw new RegistryError(
      `endpoint '${key}' is not in the registry yet. Run the REA capture workflow against the Teams web client and add the endpoint to data/endpoints.json.`,
    );
  }
  return { ...spec, url: expandUrl(spec.url) };
}

export function endpointKeys(reg: EndpointRegistry): string[] {
  return Object.keys(reg.endpoints).sort();
}
