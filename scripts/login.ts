import { parseArgs } from "node:util";
import { sessionFile } from "../src/config.js";
import { browserLogin } from "../src/login.js";

// a flag instead of TEAMS_BROWSER=msedge, which only works in sh-like shells, not in cmd or PowerShell
const { values } = parseArgs({ options: { browser: { type: "string" } } });
if (values.browser) process.env.TEAMS_BROWSER = values.browser;

console.log(`${process.env.TEAMS_BROWSER === "msedge" ? "Edge" : "Chrome"} opens Teams. Log in there and answer "Stay signed in?" with yes; the window closes by itself once you are in.`);
browserLogin(sessionFile())
  .then((r) => console.log(`saved tokens for ${r.tokenAudiences.join(", ")} -> ${r.out}`))
  .catch((e) => {
    console.error(String(e?.stack ?? e));
    process.exit(1);
  });
