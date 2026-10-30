import { sessionFile } from "../src/config.js";
import { browserLogin } from "../src/login.js";

console.log("Chrome opens Teams. Log in there; the window closes by itself once you are in.");
browserLogin(sessionFile())
  .then((r) => console.log(`saved tokens for ${r.tokenAudiences.join(", ")} -> ${r.out}`))
  .catch((e) => {
    console.error(String(e?.stack ?? e));
    process.exit(1);
  });
