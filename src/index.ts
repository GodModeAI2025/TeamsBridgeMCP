import { parseArgs } from "node:util";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { App } from "./app.js";
import { registerTools } from "./tools.js";

// --onedrive-host x sets TEAMS_ONEDRIVE_HOST=x; flags win over the environment
const flags = ["browser", "onedrive-host", "region", "partition", "chat-files-folder", "session-file", "endpoints-file"];
const { values } = parseArgs({ options: Object.fromEntries(flags.map((f) => [f, { type: "string" as const }])) });
for (const [k, v] of Object.entries(values)) {
  if (typeof v === "string") process.env["TEAMS_" + k.toUpperCase().replace(/-/g, "_")] = v;
}

const server = new McpServer({ name: "teams-bridge-mcp", version: "0.1.0" });
const app = new App();
registerTools(server, app);

const transport = new StdioServerTransport();
await server.connect(transport);
