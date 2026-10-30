// Teams messages are HTML (messagetype RichText/Html). Agents write markdown, so convert a small subset and escape the rest.

export interface Person {
  objectId: string;
  displayName: string;
}

export interface Mention {
  "@type": "http://schema.skype.com/Mention";
  itemid: string;
  mri: string;
  mentionType: "person";
  displayName: string;
}

// "@name@contoso.com" in the text mentions that person
export const MENTION_RE = /@([\w.+-]+@[\w-]+(?:\.[\w-]+)+)/g;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function inline(text: string, stash: string[]): string {
  // code, links and mentions are stashed first so emphasis markers inside them stay literal
  const keep = (html: string) => `\u0001${stash.push(html) - 1}\u0001`;
  let s = esc(text)
    .replace(/`([^`]+)`/g, (_, c) => keep(`<code>${c}</code>`))
    .replace(/\[([^\]]+)\]\(((?:https?:\/\/|mailto:)[^)\s]+)\)/g, (_, t, u) => keep(`<a href="${u}">${t}</a>`));
  s = s
    .replace(/\*\*(.+?)\*\*/g, "<b>$1</b>")
    .replace(/~~(.+?)~~/g, "<s>$1</s>")
    .replace(/(^|[^\w*])\*(?!\s)(.+?)\*(?!\w)/g, "$1<i>$2</i>")
    .replace(/(^|[^\w])_(?!\s)(.+?)_(?!\w)/g, "$1<i>$2</i>");
  return s;
}

/** Markdown subset to Teams HTML: **bold**, *italic*, ~~strike~~, `code`, [link](https://...), - and 1. lists, paragraphs, line breaks. */
export function toTeamsHtml(text: string, people: Map<string, Person> = new Map()): { content: string; mentions: Mention[] } {
  const stash: string[] = [];
  const mentions: Mention[] = [];
  const withMentions = text.replace(MENTION_RE, (all, email: string) => {
    const p = people.get(email.toLowerCase());
    if (!p) return all;
    const itemid = String(mentions.length);
    mentions.push({ "@type": "http://schema.skype.com/Mention", itemid, mri: `8:orgid:${p.objectId}`, mentionType: "person", displayName: p.displayName });
    return `\u0001${stash.push(`<span itemscope itemtype="http://schema.skype.com/Mention" itemid="${itemid}">${esc(p.displayName)}</span>`) - 1}\u0001`;
  });

  const blocks = withMentions.replace(/\r\n/g, "\n").split(/\n{2,}/);
  const html = blocks
    .map((block) => {
      let out = "";
      let list: "ul" | "ol" | undefined;
      const lines: string[] = [];
      const flush = () => {
        if (lines.length) out += `<p>${lines.join("<br>")}</p>`;
        lines.length = 0;
      };
      for (const line of block.split("\n")) {
        const m = line.match(/^\s*(?:([-*])|\d+[.)])\s+(.*)$/);
        const kind = m ? (m[1] ? "ul" : "ol") : undefined;
        if (kind !== list) {
          if (list) out += `</${list}>`;
          flush();
          if (kind) out += `<${kind}>`;
          list = kind;
        }
        if (kind) out += `<li>${inline(m![2], stash)}</li>`;
        else lines.push(inline(line, stash));
      }
      if (list) out += `</${list}>`;
      flush();
      return out;
    })
    .join("");
  return { content: html.replace(/\u0001(\d+)\u0001/g, (_, i) => stash[Number(i)]), mentions };
}

/** Teams message HTML to plain text, keeping paragraph and list breaks. */
export function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>|<\/(p|div|li)>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
