import fs from "node:fs";
import { registryFile, sessionFile } from "./config.js";
import { loadRegistry, setDetected } from "./registry.js";
import { loadSession, SessionError, SessionState } from "./session.js";
import { EndpointRegistry } from "./types.js";

function mtime(file: string): number {
  try {
    return fs.statSync(file).mtimeMs;
  } catch {
    return -1;
  }
}

export class App {
  private session?: { state: SessionState; mtime: number };
  private registry?: { reg: EndpointRegistry; mtime: number };

  // Reload when the file changed on disk, so a new login or registry edit works without restarting the server.
  getSession(): SessionState {
    const m = mtime(sessionFile());
    if (this.session?.mtime !== m) {
      this.session = { state: loadSession(sessionFile()), mtime: m };
      setDetected(this.session.state.detected);
    }
    return this.session.state;
  }

  resetSession(): void {
    this.session = undefined;
  }

  getRegistry(): EndpointRegistry {
    // endpoint URLs are expanded with the session's detected tenant values, so load the session first if there is one
    try {
      this.getSession();
    } catch {
      // no session yet: defaults and flags apply
    }
    const m = mtime(registryFile());
    if (this.registry?.mtime !== m) {
      this.registry = { reg: loadRegistry(registryFile()), mtime: m };
    }
    return this.registry.reg;
  }

  sessionStatus(): Record<string, unknown> {
    const f = sessionFile();
    if (!fs.existsSync(f)) {
      return { present: false, path: f };
    }
    try {
      const s = loadSession(f);
      return {
        present: true,
        path: f,
        cookieCount: s.cookies.length,
        cookieNames: s.cookieNames(),
        tokenAudiences: s.tokenAudiences(),
        detected: s.detected,
      };
    } catch (e) {
      return { present: true, path: f, error: (e as Error).message };
    }
  }
}
