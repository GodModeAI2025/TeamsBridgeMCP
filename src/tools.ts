import fs from "node:fs";
import path from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { App } from "./app.js";
import { apiRequest, ApiError } from "./http.js";
import { endpoint } from "./registry.js";
import { sessionFile } from "./config.js";
import { browserLogin } from "./login.js";
import { ApiResponse } from "./types.js";
import { SessionError } from "./session.js";
import { htmlToText, MENTION_RE, Person, toTeamsHtml } from "./format.js";

function ok(data: unknown) {
  const text = typeof data === "string" ? data : JSON.stringify(data, null, 2);
  return { content: [{ type: "text" as const, text }] };
}

function fail(e: unknown) {
  const msg = e instanceof Error ? e.message : String(e);
  return { content: [{ type: "text" as const, text: `error: ${msg}` }], isError: true };
}

function payload(res: ApiResponse): unknown {
  return res.buffer ? { bytes: res.buffer.length } : res.json ?? res.text;
}

export function registerTools(server: McpServer, app: App): void {
  let loginRunning: Promise<unknown> | undefined;
  let lastLogin: string | undefined;
  let refreshing: Promise<void> | undefined;

  // Visible Chrome login in the background; MCP clients with short tool timeouts must not wait for the user's password.
  function startVisibleLogin(): void {
    if (loginRunning) return;
    loginRunning = (refreshing ?? Promise.resolve())
      .catch(() => {})
      .then(() => browserLogin(sessionFile()))
      .then((r) => {
        app.resetSession();
        lastLogin = `ok: tokens for ${r.tokenAudiences.length} APIs saved`;
      })
      .catch((e) => {
        lastLogin = `failed: ${e instanceof Error ? e.message : String(e)}`;
      })
      .finally(() => {
        loginRunning = undefined;
      });
  }

  // Silent re-login with the saved browser profile; only if Microsoft wants a password/MFA the visible window opens.
  async function refreshLogin(): Promise<void> {
    if (loginRunning) throw new Error("Teams sign-in is in progress in the Chrome window; retry once it has closed");
    refreshing ??= browserLogin(sessionFile(), { headless: true, timeoutMs: 60_000 })
      .then(() => app.resetSession())
      .finally(() => {
        refreshing = undefined;
      });
    try {
      await refreshing;
    } catch {
      startVisibleLogin();
      throw new Error("Teams session expired and needs a sign-in: a Chrome window was opened, sign in there and retry");
    }
  }

  function expiresSoon(token: string): boolean {
    try {
      const exp = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8")).exp;
      return typeof exp === "number" && exp * 1000 < Date.now() + 60_000;
    } catch {
      return false;
    }
  }

  // Every API call goes through here: refreshes a missing/expiring token up front and retries once on HTTP 401.
  async function call(key: string, opts: Parameters<typeof apiRequest>[2] = {}): Promise<ApiResponse> {
    const spec = endpoint(app.getRegistry(), key);
    if (spec.tokenSource) {
      let token: string | undefined;
      try {
        token = app.getSession().token(spec.tokenSource);
      } catch (e) {
        if (!(e instanceof SessionError)) throw e;
      }
      if (!token || expiresSoon(token)) await refreshLogin();
    }
    try {
      return await apiRequest(spec, app.getSession(), opts);
    } catch (e) {
      if (!(e instanceof ApiError) || e.status !== 401) throw e;
      await refreshLogin();
      return apiRequest(spec, app.getSession(), opts);
    }
  }
  // properties.files is a JSON string; objectUrl is what download_file takes as fileReference
  function messageFiles(m: any): { name: string; type: string; fileReference: string }[] | undefined {
    try {
      const files = JSON.parse(m.properties?.files ?? "[]") as any[];
      if (!files.length) return undefined;
      return files.map((f) => ({ name: f.fileName ?? f.title, type: f.fileType ?? f.type, fileReference: f.objectUrl }));
    } catch {
      return undefined;
    }
  }

  function myObjectId(): string | undefined {
    const token = app.getSession().token(endpoint(app.getRegistry(), "chats.list").tokenSource ?? "");
    return token ? JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8")).oid : undefined;
  }

  // A host that times out retries the call; the same send within 30 s gets the first result instead of a second post.
  // Not for edit_message: an idempotent PUT must go out again, or A -> B -> A would report success without sending.
  const recent = new Map<string, Promise<unknown>>();
  function singleFlight<T>(key: string, fn: () => Promise<T>): Promise<T> {
    let p = recent.get(key) as Promise<T> | undefined;
    if (!p) {
      p = fn();
      recent.set(key, p);
      p.then(
        () => setTimeout(() => recent.delete(key), 30_000).unref(),
        () => recent.delete(key),
      );
    }
    return p;
  }

  // Teams localizes the OneDrive folder for chat uploads; uploading to the wrong name creates a second one.
  // After a language switch both can exist, and the one Teams writes to now is the one changed last.
  const CHAT_FILES_FOLDERS = ["Microsoft Teams Chat Files", "Microsoft Teams-Chatdateien"];
  let chatFilesFolder: string | undefined;
  async function findChatFilesFolder(): Promise<string> {
    if (process.env.TEAMS_CHAT_FILES_FOLDER) return process.env.TEAMS_CHAT_FILES_FOLDER;
    if (chatFilesFolder) return chatFilesFolder;
    let best: { folder: string; modified: string } | undefined;
    for (const folder of CHAT_FILES_FOLDERS) {
      try {
        const modified = String(((await call("files.folder", { params: { folder } })).json as any)?.lastModifiedDateTime ?? "");
        if (!best || modified > best.modified) best = { folder, modified };
      } catch (e) {
        if (!(e instanceof ApiError)) throw e;
        if (e.status === 404) continue;
        // 401/403 is a sign-in or access problem, not a missing folder
        if (e.status === 401 || e.status === 403) throw e;
        return best?.folder ?? CHAT_FILES_FOLDERS[0]; // unclear answer: use what we have, but ask again next time
      }
    }
    // none yet: the upload creates the default one
    return best ? (chatFilesFolder = best.folder) : CHAT_FILES_FOLDERS[0];
  }

  // chatsvc message format, as sent by the Teams web client
  function postMessage(conversationId: string, key: string, msg: Record<string, unknown>, replyToMessageId?: string) {
    return call(key, {
      params: { conversationId, replyToMessageId },
      body: { contenttype: "text", clientmessageid: String(Date.now()) + String(Math.floor(Math.random() * 1e6)), ...msg },
    });
  }

  // markdown subset -> Teams HTML; "@name@contoso.com" becomes a mention of that person
  async function htmlBody(text: string): Promise<Record<string, unknown>> {
    const emails = [...new Set([...text.matchAll(MENTION_RE)].map((m) => m[1].toLowerCase()))];
    const people = new Map<string, Person>();
    if (emails.length) {
      const profiles = ((await call("users.byMail", { body: emails })).json as { value?: any[] })?.value ?? [];
      for (const p of profiles) if (p.objectId) people.set(String(p.email).toLowerCase(), p);
      const missing = emails.filter((e) => !people.has(e));
      if (missing.length) throw new Error(`unknown user(s) to mention: ${missing.join(", ")}`);
    }
    const { content, mentions } = toTeamsHtml(text, people);
    return { content, messagetype: "RichText/Html", ...(mentions.length ? { properties: { mentions: JSON.stringify(mentions) } } : {}) };
  }

  // 1:1 chats have no topic; Teams shows the other person's name, encoded as the second GUID in the chat ID
  async function oneOnOneNames(chatIds: string[]): Promise<Map<string, string>> {
    const names = new Map<string, string>();
    try {
      const reg = app.getRegistry();
      const me = myObjectId();
      const partner = new Map<string, string>();
      for (const id of chatIds) {
        const m = id.match(/^19:([0-9a-f-]{36})_([0-9a-f-]{36})@unq\.gbl\.spaces$/);
        if (m && me) partner.set(id, m[1] === me ? m[2] : m[1]);
      }
      if (partner.size === 0) return names;
      const res = await call("users.profiles", {
        body: [...new Set(partner.values())].map((g) => `8:orgid:${g}`),
      });
      const byId = new Map<string, string>();
      for (const p of (res.json as { value?: any[] })?.value ?? []) byId.set(p.objectId, p.displayName);
      for (const [id, g] of partner) if (byId.has(g)) names.set(id, byId.get(g)!);
    } catch {
      // names are a convenience; the chat list is still useful without them
    }
    return names;
  }

  server.tool(
    "teams_status",
    "Check Teams session export and endpoint registry status",
    {},
    async () => {
      try {
        const reg = app.getRegistry();
        return ok({
          session: app.sessionStatus(),
          login: loginRunning ? "in progress" : lastLogin ?? "not started",
          registry: {
            endpointCount: Object.keys(reg.endpoints).length,
            endpoints: Object.keys(reg.endpoints).sort(),
          },
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "teams_login",
    "Open Chrome on Teams so the user can sign in; the session is saved automatically. Returns immediately, check teams_status afterwards.",
    {},
    async () => {
      if (loginRunning) return ok("login already in progress: finish signing in in the Chrome window");
      startVisibleLogin();
      return ok("Chrome window opened: sign in to Teams there. It closes by itself when done; then call teams_status.");
    },
  );

  server.tool(
    "list_chats",
    "List recent conversations (chats, meeting chats, channels) of the signed-in user, newest activity first",
    { limit: z.number().int().min(1).max(200).optional().describe("Max conversations (default 50)") },
    async ({ limit }) => {
      try {
        const res = await call("chats.list", {
          params: { pageSize: limit },
        });
        // raw chatsvc conversations are ~3 KB each; return only what an agent needs to pick one
        const convs = (res.json as { conversations?: any[] })?.conversations;
        if (!convs) return ok(payload(res));
        const names = await oneOnOneNames(convs.map((c) => c.id));
        return ok(
          convs.map((c) => ({
            id: c.id,
            topic: c.threadProperties?.topic ?? names.get(c.id) ?? (c.id === "48:notes" ? "Notes to self" : undefined),
            type: c.threadProperties?.threadType ?? c.type,
            lastMessageAt: c.lastMessage?.composetime,
            lastFrom: c.lastMessage?.imdisplayname,
            lastText: htmlToText(String(c.lastMessage?.content ?? "")).slice(0, 120),
          })),
        );
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "list_channels",
    "List your teams and their channels (channel id works with list_messages and send_message)",
    { teamId: z.string().optional().describe("Only this team (id from a previous call)") },
    async ({ teamId }) => {
      try {
        const res = await call("channels.list");
        const teams = ((res.json as { teams?: any[] })?.teams ?? []).filter((t) => !teamId || t.id === teamId);
        return ok(
          teams.map((t) => ({
            id: t.id,
            name: t.displayName,
            channels: (t.channels ?? []).map((c: any) => ({ id: c.id, name: c.displayName })),
          })),
        );
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "list_messages",
    "List messages of a channel or chat conversation",
    {
      conversationId: z.string().describe("Channel or chat conversation ID"),
      kind: z.enum(["channel", "chat"]).default("channel"),
      cursor: z.string().optional().describe("olderCursor from the previous result, to get the next older page"),
      limit: z.number().int().min(1).max(200).optional(),
    },
    async ({ conversationId, kind, cursor, limit }) => {
      try {
        // the cursor is the query of chatsvc's backwardLink; only these two values are taken from it, never a host or path
        const page = new URLSearchParams(cursor ?? "");
        const res = await call("messages.list", {
          params: { conversationId, pageSize: limit, startTime: page.get("startTime") ?? undefined, syncState: page.get("syncState") ?? undefined },
        });
        // drop call/membership events and HTML; raw pages are ~100 KB per 50 messages
        const json = res.json as { messages?: any[]; _metadata?: { backwardLink?: string } };
        const msgs = json?.messages;
        if (!msgs) return ok(payload(res));
        const back = json._metadata?.backwardLink ? new URL(json._metadata.backwardLink).search.slice(1) : undefined;
        return ok({
          messages: msgs
            .filter((m) => m.messagetype === "Text" || String(m.messagetype).startsWith("RichText"))
            .map((m) => ({
              id: m.id,
              from: m.imdisplayname,
              at: m.composetime,
              text: htmlToText(String(m.content ?? "")).slice(0, 2000),
              files: messageFiles(m),
            })),
          // chatsvc repeats the same link at the end of the history
          olderCursor: msgs.length && back !== cursor ? back : undefined,
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "send_message",
    "Send a message to a chat or channel; in channels pass replyToMessageId (the thread's root post id) to reply in that thread. Returns the new messageId (for edit_message/react_message). The same text to the same conversation within 30 s is sent only once.",
    {
      conversationId: z.string().describe("Conversation ID from list_chats, list_channels or start_chat"),
      text: z
        .string()
        .min(1)
        .describe("Message text. Markdown subset: **bold**, *italic*, ~~strike~~, `code`, [text](https://...), - lists (numbered lines stay plain text). Mention a person with @their-email, e.g. @alex@contoso.com"),
      kind: z.enum(["channel", "chat"]).default("channel"),
      replyToMessageId: z.string().optional(),
    },
    async ({ conversationId, text, kind, replyToMessageId }) => {
      try {
        const key = replyToMessageId ? "messages.reply" : "messages.send";
        const res = await singleFlight(JSON.stringify([key, conversationId, replyToMessageId, text]), async () =>
          postMessage(conversationId, key, await htmlBody(text), replyToMessageId),
        );
        return ok({ messageId: res.headers.get("location")?.split("/").pop(), ...(payload(res) as object) });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "edit_message",
    "Replace the text of one of your own messages (messageId from send_message or list_messages)",
    {
      conversationId: z.string(),
      messageId: z.string(),
      text: z.string().min(1).describe("New text, same markdown and @email mentions as send_message"),
    },
    async ({ conversationId, messageId, text }) => {
      try {
        // PUT replaces the whole message: carry over attachments, or they disappear
        const current = (await call("messages.get", { params: { conversationId, messageId } })).json as any;
        const body = { contenttype: "text", ...(await htmlBody(text)), id: messageId, type: "Message", conversationid: conversationId } as any;
        if (current?.properties?.files) body.properties = { ...body.properties, files: current.properties.files };
        const res = await call("messages.edit", { params: { conversationId, messageId }, body });
        return ok({ edited: messageId, status: res.status });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "react_message",
    "Add a reaction to a message",
    {
      conversationId: z.string(),
      messageId: z.string(),
      reaction: z.enum(["like", "heart", "laugh", "surprised", "sad", "angry"]).default("like"),
    },
    async ({ conversationId, messageId, reaction }) => {
      try {
        const res = await call("messages.react", {
          params: { conversationId, messageId },
          body: { emotions: { key: reaction, value: Date.now() } },
        });
        return ok({ reacted: reaction, status: res.status });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "search_messages",
    "Search messages across chats and channels",
    {
      query: z.string().min(1),
      limit: z.number().int().min(1).max(100).optional(),
    },
    async ({ query, limit }) => {
      try {
        // the service rejects requests without scenario/cvid/logicalId
        const res = await call("search.messages", {
          body: {
            EntityRequests: [
              {
                entityType: "Message",
                contentSources: ["Teams"],
                propertySet: "Optimized",
                query: { queryString: query },
                from: 0,
                size: limit ?? 25,
              },
            ],
            cvid: crypto.randomUUID(),
            logicalId: crypto.randomUUID(),
            scenario: { Name: "powerbar" },
          },
        });
        const results = (res.json as any)?.EntitySets?.[0]?.ResultSets?.[0]?.Results ?? [];
        return ok(
          results.map((r: any) => ({
            conversationId: r.Source?.ClientThreadId,
            messageId: String(r.Source?.ClientConversationId ?? "").split(";messageid=")[1],
            from: r.Source?.From?.EmailAddress?.Name,
            at: r.Source?.DateTimeSent,
            text: r.Source?.Preview || r.HitHighlightedSummary,
          })),
        );
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "upload_file",
    "Send a local file to a chat: uploads it to your OneDrive Teams chat files folder, gives the chat members read access and posts it as a message (like Teams does). Use conversationId '48:notes' to send to yourself.",
    {
      conversationId: z.string().describe("Chat conversation ID from list_chats"),
      path: z.string().describe("Local file path to upload"),
      name: z.string().optional().describe("File name in Teams (defaults to the local file name)"),
      text: z.string().optional().describe("Optional message text shown with the file"),
    },
    async ({ conversationId, path: filePath, name, text }) => {
      try {
        const abs = path.resolve(filePath);
        return ok(await singleFlight(JSON.stringify(["upload", conversationId, abs, name, text]), () => uploadAndPost(abs)));
      } catch (e) {
        return fail(e);
      }

      async function uploadAndPost(abs: string) {
        if (!fs.existsSync(abs)) throw new Error(`file not found: ${abs}`);
        // shortcut: single PUT upload, SharePoint caps it at 250 MB; switch to an upload session for larger files
        const up = await call("files.upload", {
          params: { folder: await findChatFilesFolder(), name: name ?? path.basename(abs) },
          body: fs.readFileSync(abs),
          timeoutMs: 300000,
        });
        const itemId = (up.json as { id: string }).id;
        const item = (await call("files.item", { params: { itemId } })).json as any;

        // Notes-to-self needs no sharing; other chats: read access for every person in the chat
        let sharedWith = 0;
        if (conversationId !== "48:notes") {
          const thread = (await call("threads.get", { params: { conversationId } })).json as any;
          const me = myObjectId();
          const people = ((thread?.members ?? []) as any[])
            .map((m) => String(m.id).match(/^8:orgid:([0-9a-f-]{36})$/)?.[1])
            .filter((g): g is string => !!g && g !== me);
          if (people.length) {
            await call("files.invite", {
              params: { itemId },
              body: {
                recipients: people.map((objectId) => ({ objectId })),
                roles: ["read"],
                requireSignIn: true,
                sendInvitation: false,
              },
            });
          }
          sharedWith = people.length;
        }

        const uid = item.sharepointIds.listItemUniqueId;
        const objectUrl = decodeURI(item.webUrl);
        const baseUrl = item.sharepointIds.siteUrl + "/";
        const ext = path.extname(item.name).slice(1);
        const file = {
          "@type": "http://schema.skype.com/File",
          version: 2,
          id: uid,
          itemid: uid,
          baseUrl,
          objectUrl,
          type: ext,
          title: item.name,
          fileName: item.name,
          fileType: ext,
          state: "active",
          fileInfo: { itemId: null, fileUrl: objectUrl, siteUrl: baseUrl, serverRelativeUrl: "", shareUrl: null, shareId: null },
          fileChicletState: { serviceName: "p2p", state: "active" },
          botFileProperties: {},
          permissionScope: "users",
          sharepointIds: { driveId: null, listId: null, listItemUniqueId: uid, siteId: null, siteUrl: null, webId: null },
        };
        const res = await postMessage(conversationId, "messages.send", {
          content: text ?? "",
          messagetype: "RichText/Html",
          properties: { files: JSON.stringify([file]) },
        });
        return { sent: payload(res), fileName: item.name, fileReference: objectUrl, sharedWith };
      }
    },
  );

  server.tool(
    "download_file",
    "Download a file from a conversation to a local path",
    {
      fileReference: z.string().url().describe("fileReference (SharePoint/OneDrive URL) from list_messages files[]"),
      outPath: z.string().describe("Local output path"),
    },
    async ({ fileReference, outPath }) => {
      try {
        const ref = new URL(fileReference);
        // the SharePoint token goes to this host, so only SharePoint hosts are allowed
        if (ref.protocol !== "https:" || !ref.hostname.endsWith(".sharepoint.com")) {
          throw new Error(`not a SharePoint/OneDrive file: ${ref.host}`);
        }
        const res = await call("files.download", {
          params: { host: ref.host, shareId: "u!" + Buffer.from(fileReference).toString("base64url") },
          raw: true,
          timeoutMs: 300000,
          maxBytes: 250 * 1024 * 1024,
        });
        const buf = res.buffer;
        if (!buf) throw new Error("download returned no data");
        const abs = path.resolve(outPath);
        fs.mkdirSync(path.dirname(abs), { recursive: true });
        fs.writeFileSync(abs, buf);
        return ok({ savedTo: abs, bytes: buf.length });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "start_chat",
    "Start (or reopen) a chat with people by e-mail; returns the conversationId for send_message/upload_file. One person without topic = 1:1 chat (existing one is reused), otherwise a new group chat.",
    {
      emails: z.array(z.string().email()).min(1).describe("E-mail addresses of the other participants"),
      topic: z.string().optional().describe("Group chat name"),
    },
    async ({ emails, topic }) => {
      try {
        const profiles = ((await call("users.byMail", { body: emails })).json as { value?: any[] })?.value ?? [];
        const ids = profiles.map((p) => p.objectId).filter(Boolean);
        const missing = emails.filter((e) => !profiles.some((p) => String(p.email).toLowerCase() === e.toLowerCase()));
        if (ids.length !== emails.length) throw new Error(`unknown user(s): ${missing.join(", ") || emails.join(", ")}`);
        const me = myObjectId();
        const oneOnOne = ids.length === 1 && !topic;
        const res = await call("threads.create", {
          body: {
            members: [me, ...ids].map((g) => ({ id: `8:orgid:${g}`, role: "Admin" })),
            properties: oneOnOne
              ? { threadType: "chat", fixedRoster: "true", uniquerosterthread: "true" }
              : { threadType: "chat", ...(topic ? { topic } : {}) },
          },
        });
        const conversationId = decodeURIComponent(res.headers.get("location")?.split("/threads/")[1] ?? "");
        if (!conversationId) throw new Error("chat created but no conversation id returned");
        return ok({ conversationId, members: profiles.map((p) => p.displayName) });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.tool(
    "list_meetings",
    "List meetings in your calendar from now on",
    { days: z.number().int().min(1).max(60).optional().describe("How many days ahead (default 7)") },
    async ({ days }) => {
      try {
        const start = new Date();
        const end = new Date(start.getTime() + (days ?? 7) * 86_400_000);
        const res = await call("meetings.list", { params: { StartDate: start.toISOString(), EndDate: end.toISOString() } });
        return ok(
          ((res.json as { value?: any[] })?.value ?? []).map((m) => ({
            subject: m.subject,
            start: m.startTime,
            end: m.endTime,
            organizer: m.organizerName,
            location: m.location,
            joinUrl: m.skypeTeamsMeetingUrl || m.joinOnlineMeetingUrl || undefined,
            response: m.myResponseType,
          })),
        );
      } catch (e) {
        return fail(e);
      }
    },
  );
}
