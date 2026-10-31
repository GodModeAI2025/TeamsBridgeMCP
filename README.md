# TeamsBridgeMCP

An MCP server for Microsoft Teams that talks to the same internal REST endpoints the Teams web client (teams.microsoft.com) uses. It does **not** use Microsoft Graph, so you need no Azure app registration and no admin consent. You sign in once in Chrome with your normal account and the server reuses that session.

All endpoints live in one registry file (`data/endpoints.json`). When Microsoft changes something, you patch that file, not the code.

Landing page with architecture diagram: https://godmodeai2025.github.io/TeamsBridgeMCP/

> Independent open-source project, not affiliated with, sponsored or endorsed by Microsoft. See [Trademarks](#trademarks). The endpoints used here are undocumented and can change at any time. Check your organization's policies before using this with a work account.

## Tools

| Tool | What it does |
|---|---|
| `teams_status` | Session and registry status |
| `teams_login` | Opens Chrome so you can sign in; session is saved automatically |
| `list_chats` | Recent chats, meeting chats and channels, newest first |
| `list_channels` | Your teams and their channels |
| `list_messages` | Messages of a chat or channel, including attached files; `olderCursor` fetches the next older page |
| `send_message` | Send to a chat or channel (the same text to the same chat within 30 s is sent once, so a client retry does not post twice); `replyToMessageId` replies in a channel thread. Markdown (bold, italic, strike, code, links, lists) becomes Teams formatting, `@name@contoso.com` mentions that person. Returns the `messageId` |
| `edit_message` | Change the text of one of your messages, attachments stay |
| `react_message` | Add a reaction (like, heart, laugh, surprised, sad, angry) |
| `start_chat` | Start or reopen a chat by e-mail address |
| `search_messages` | Search across chats and channels |
| `upload_file` | Upload a local file to your OneDrive and post it into a chat |
| `download_file` | Download a file from a message (SharePoint or OneDrive links only, up to 250 MB) |
| `list_meetings` | Upcoming meetings from your calendar |

## Requirements

- Node.js 20 or newer
- Google Chrome or Microsoft Edge installed (`playwright-core` drives your local browser and does not download one)
- A Microsoft 365 account with Teams

## Setup

```bash
git clone https://github.com/GodModeAI2025/TeamsBridgeMCP.git
cd TeamsBridgeMCP
npm install
npm run build
npm run login        # Chrome opens, sign in, the window closes by itself
```

Prefer Edge? `npm run login -- --browser msedge`, and pass `--browser msedge` to the server. The commands are the same on macOS and Windows (Terminal, cmd or PowerShell).

`npm run login` writes `data/session.json` (gitignored; mode 0600 on macOS and Linux, on Windows it is protected by the folder's permissions, so keep the project inside your user profile) and keeps a browser profile in `data/browser-profile/` so later sign-ins usually need no password. You can also call the `teams_login` tool from your agent instead.

Expired tokens (chat about 24 h, files about 1 h) are refreshed by the server itself with a headless sign-in. Only when Microsoft asks for password or MFA does it open the Chrome window again. If Microsoft asks "Stay signed in?" during the first sign-in, choose yes: on a Mac without company device sign-in the silent refresh only works then, otherwise the stored profile keeps no login. HTTP 403 is passed through unchanged: it means no access to that chat or a wrong region, and a new sign-in would not help.

## Configuration

Usually nothing to configure. At sign-in the server reads the chat region, the partition and your OneDrive host from the Teams client's own requests and stores them with the session; `teams_status` shows them. Flags or environment variables override the detected values, flags win over the environment.

| Flag | Variable | Default | When to change it |
|---|---|---|---|
| `--onedrive-host` | `TEAMS_ONEDRIVE_HOST` | detected at sign-in | Your OneDrive host for `upload_file`, e.g. `contoso-my.sharepoint.com`. Set it only if detection fails: open OneDrive in the browser and copy the host from the address bar. |
| `--region` | `TEAMS_REGION` | detected, else `emea` | Chat service region of your tenant, e.g. `emea`, `amer` or `apac`. |
| `--partition` | `TEAMS_PARTITION` | detected, else `emea-02` | Middle-tier partition used for profiles and meetings. |
| `--chat-files-folder` | `TEAMS_CHAT_FILES_FOLDER` | found in your OneDrive | OneDrive folder Teams uses for chat uploads. The server looks for the English and German name and takes the one changed last; set this if Teams runs in another language. |
| `--browser` | `TEAMS_BROWSER` | `chrome` | Browser for sign-in: `chrome` or `msedge` |
| `--session-file` | `TEAMS_SESSION_FILE` | `data/session.json` | Different session location |
| `--endpoints-file` | `TEAMS_ENDPOINTS_FILE` | `data/endpoints.json` | Your own registry file |

Finding region and partition by hand is only needed after `npm run export-session` (that path does not detect them) or when every chat returns HTTP 403, which is how a wrong region shows up. Open Teams in the browser, open DevTools (Network tab), click into a chat. Requests to `.../api/chatsvc/<region>/v1/...` show the region, requests to `.../api/mt/part/<partition>/...` show the partition.

### Register with an MCP client

Claude Code:

```bash
claude mcp add teams -- node /absolute/path/TeamsBridgeMCP/dist/index.js
```

Claude Desktop (`claude_desktop_config.json`) and most other clients, here with a flag that overrides a detected value:

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": [
        "/absolute/path/TeamsBridgeMCP/dist/index.js",
        "--region", "emea"
      ]
    }
  }
}
```

The same with environment variables:

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": ["/absolute/path/TeamsBridgeMCP/dist/index.js"],
      "env": { "TEAMS_BROWSER": "msedge" }
    }
  }
}
```

On Windows the path in `args` needs forward slashes or doubled backslashes, e.g. `"C:/Users/you/TeamsBridgeMCP/dist/index.js"`.

An unknown flag stops the server with an error, so typos show up right away.

## Test

```bash
npm test             # builds, checks URL expansion, starts the server and calls teams_status
```

This needs no Teams session.

## When an endpoint breaks: capture workflow

Tools report `HTTP 404/410` or auth errors when Microsoft moved an endpoint. Re-capture it:

1. Start Chrome with a debug port and a separate profile (current Chrome ignores the port on the default profile):
   ```bash
   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
     --remote-debugging-port=9222 --user-data-dir="$HOME/.teams-cdp-profile"
   ```
   On Windows (PowerShell):
   ```powershell
   & "$env:ProgramFiles\Google\Chrome\Application\chrome.exe" `
     --remote-debugging-port=9222 --user-data-dir="$env:USERPROFILE\.teams-cdp-profile"
   ```
   Edge works the same, e.g. `& "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"`.
   Open https://teams.microsoft.com there and sign in.
2. `npm run capture -- --label messages`
3. Do the action in Teams (open a channel, send a message, page back).
4. After each action set a marker: `npm run mark -- "sent message in channel X"`
5. Stop with Ctrl+C. The requests are in `captures/messages-<timestamp>.jsonl`.
6. Copy the relevant request into `data/endpoints.json` and call the tool again.

`npm run export-session` exports the cookies of that debug browser to `data/session.json` if you prefer it over `npm run login`.

### Registry format

```json
{
  "version": 1,
  "updated": "2026-10-29T14:00:00.000Z",
  "endpoints": {
    "messages.send": {
      "method": "POST",
      "url": "https://teams.cloud.microsoft/api/chatsvc/${TEAMS_REGION}/v1/users/ME/conversations/{conversationId}/messages",
      "auth": "bearer",
      "tokenSource": "https://ic3.teams.office.com",
      "bodyType": "json",
      "params": [{ "name": "conversationId", "in": "path", "required": true }]
    }
  }
}
```

- `${NAME}` in `url` is replaced by the environment variable (see Configuration)
- `{name}` in `url` is a path parameter; all other parameters go into the query string
- `auth`: `cookie`, `bearer`, `cookie+bearer` or `none`
- `bearer` needs `tokenSource`: the audience of a token captured at sign-in (e.g. `https://ic3.teams.office.com`) or a cookie name
- `headers` values containing `{uuid}` get a fresh UUID per request

## Security

- `data/session.json` holds your full Teams session. Never commit or share it. If it leaks, sign out of Teams everywhere or reset your password.
- `captures/` contains request headers and bodies, including tokens. Gitignored, do not share.
- The server can do exactly what your signed-in account can do, nothing more.
- Messages from other people are input your agent reads. Keep tool approval on in your MCP client for `send_message`, `edit_message`, `react_message`, `upload_file` and `start_chat`, so a crafted message cannot make the agent post or share files unseen.

## Limits

- Undocumented endpoints, no support from Microsoft. Use at your own risk.
- No real-time events (push/WebSocket); reads poll.
- Meetings: list and details yes, joining or controlling no.
- `upload_file` uses a single PUT, which SharePoint caps at 250 MB.

## Legal background

This section explains why the author considers publishing this project lawful. It is provided for transparency only and is not legal advice.

- **Interfaces are not protected as such.** Under EU law, the functionality of a computer program, its programming language and its data formats are not protected by copyright (Court of Justice of the European Union, judgment of 2 May 2012, SAS Institute, C-406/10). This project describes and uses interfaces. It does not copy any part of Microsoft's software.
- **Observing a program is permitted.** A lawful user may observe, study and test the functioning of a program in order to determine the ideas and principles behind it, while running it as intended (Art. 5(3) Directive 2009/24/EC, implemented in Germany as § 69d(3) UrhG). The endpoints were identified by observing the network traffic of the Teams web client during normal, authorized use. No code was decompiled or disassembled.
- **Reverse engineering of available products is permitted.** Acquiring information by observing, studying or testing a product that has been made available to the public is lawful (Art. 3(1)(b) Directive (EU) 2016/943, implemented in Germany as § 3(1) No. 2 GeschGehG).
- **No protection measures are circumvented.** The software only uses the session the user establishes by signing in with their own account in a regular browser. It grants no additional rights and works only with the permissions that account already has.
- **Contractual terms remain the user's responsibility.** Terms of service of Microsoft and the policies of the organization that manages a tenant may restrict automated access. They bind the user of an account, not the publication of this source code. Each user is responsible for checking that their use complies with the terms and policies that apply to them.

## Private project

TeamsBridgeMCP is a private project. The author created it entirely in their own time, with their own equipment and resources, on their own initiative. It was not created in the course of any employment, as part of any job duties, or on behalf of or commissioned by any employer, client or other organization. It contains no code, data, documents or confidential information belonging to any employer or client. No employer or organization owns, maintains, endorses or is responsible for this project.

## Responsible use

- Use it only with an account you are allowed to automate. On a work or school tenant, check your organization's IT and acceptable-use policies first, or ask your administrator.
- Chats contain other people's messages and personal data. Do not collect, export or monitor conversations beyond what you would do by hand.
- The software comes as is, without warranty (Apache License 2.0, sections 7 and 8).

## Trademarks

Microsoft, Microsoft Teams, Microsoft 365, Microsoft Edge, Entra ID, OneDrive and SharePoint are trademarks of the Microsoft group of companies. Chrome is a trademark of Google LLC. They are used here only to describe what this software works with. This project is independent and not affiliated with, sponsored or endorsed by Microsoft or Google. It ships no Microsoft logos, icons or code. All other names belong to their respective owners.

## License

Apache License 2.0, see [LICENSE](LICENSE).

## Legal notice

Impressum: https://godmodeai2025.github.io/MarkZimmermann/
