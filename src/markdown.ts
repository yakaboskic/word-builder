import { marked, type Tokens } from "marked";
import type { ContentBuilder } from "./content-builder";
import type { RichText } from "./xml";

export interface MarkdownMeta {
  title?: string;
  date?: string;
  author?: string;
  [key: string]: string | undefined;
}

export function parseFrontmatter(md: string): { meta: MarkdownMeta; body: string } {
  const m = md.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/);
  if (!m) return { meta: {}, body: md };
  const meta: MarkdownMeta = {};
  for (const line of m[1].split(/\r?\n/)) {
    const kv = line.match(/^(\w+)\s*:\s*(.*)$/);
    if (kv) meta[kv[1]] = kv[2].trim().replace(/^["'](.*)["']$/, "$1");
  }
  return { meta, body: m[2] };
}

type InlineToken = Tokens.Text | Tokens.Strong | Tokens.Em | Tokens.Codespan | Tokens.Br | Tokens.Del | Tokens.Link | Tokens.Generic;

function inlineToRich(tokens: InlineToken[], inherit: { bold?: boolean; italic?: boolean } = {}): RichText[] {
  const out: RichText[] = [];
  for (const t of tokens) {
    const type = (t as { type: string }).type;
    if (type === "text") {
      const text = t as Tokens.Text;
      if (text.tokens && text.tokens.length > 0) {
        out.push(...inlineToRich(text.tokens as InlineToken[], inherit));
      } else {
        out.push({ text: text.text, ...inherit });
      }
    } else if (type === "strong") {
      const strong = t as Tokens.Strong;
      out.push(...inlineToRich(strong.tokens as InlineToken[], { ...inherit, bold: true }));
    } else if (type === "em") {
      const em = t as Tokens.Em;
      out.push(...inlineToRich(em.tokens as InlineToken[], { ...inherit, italic: true }));
    } else if (type === "codespan") {
      const code = t as Tokens.Codespan;
      out.push({ text: code.text, ...inherit });
    } else if (type === "br") {
      out.push({ text: "\n", ...inherit });
    } else if (type === "del") {
      const del = t as Tokens.Del;
      out.push(...inlineToRich(del.tokens as InlineToken[], inherit));
    } else if (type === "link") {
      const link = t as Tokens.Link;
      out.push(...inlineToRich(link.tokens as InlineToken[], inherit));
    } else if (type === "escape") {
      out.push({ text: (t as Tokens.Escape).text, ...inherit });
    } else if ("raw" in t && typeof (t as { raw: unknown }).raw === "string") {
      out.push({ text: (t as { raw: string }).raw, ...inherit });
    }
  }
  return out;
}

function richToString(rich: RichText[]): string {
  return rich.map((r) => r.text).join("");
}

function renderToken(token: Tokens.Generic, builder: ContentBuilder): void {
  switch (token.type) {
    case "heading": {
      const t = token as Tokens.Heading;
      const text = t.tokens ? richToString(inlineToRich(t.tokens as InlineToken[])) : t.text;
      if (t.depth === 1) builder.h1(text);
      else if (t.depth === 2) builder.h2(text);
      else if (t.depth === 3) builder.h3(text);
      else builder.label(text);
      break;
    }
    case "paragraph": {
      const t = token as Tokens.Paragraph;
      const rich = t.tokens ? inlineToRich(t.tokens as InlineToken[]) : [{ text: t.text }];
      builder.p(rich);
      break;
    }
    case "list": {
      const t = token as Tokens.List;
      for (const item of t.items) {
        const flattened = flattenListItem(item.tokens ?? []);
        const rich = flattened.length > 0
          ? inlineToRich(flattened as InlineToken[])
          : [{ text: item.text }];
        if (t.ordered) builder.numbered(rich);
        else builder.bullet(rich);
      }
      break;
    }
    case "table": {
      const t = token as Tokens.Table;
      const headers = t.header.map((cell) => cell.text);
      const rows = t.rows.map((row) => row.map((cell) => cell.text));
      builder.table(headers, rows);
      break;
    }
    case "blockquote": {
      const t = token as Tokens.Blockquote;
      const inner = (t.tokens ?? [])
        .filter((x) => x.type === "paragraph")
        .map((x) => {
          const p = x as Tokens.Paragraph;
          return p.tokens ? richToString(inlineToRich(p.tokens as InlineToken[])) : p.text;
        })
        .join(" ");
      builder.note(inner, "info", "Note: ");
      break;
    }
    case "code": {
      const t = token as Tokens.Code;
      builder.p(t.text, { font: "Consolas", size: 20 });
      break;
    }
    case "hr": {
      builder.pageBreak();
      break;
    }
    case "space":
    default:
      break;
  }
}

function flattenListItem(tokens: Tokens.Generic[]): Tokens.Generic[] {
  const out: Tokens.Generic[] = [];
  for (const t of tokens) {
    if ((t.type === "text" || t.type === "paragraph") && "tokens" in t && Array.isArray(t.tokens)) {
      out.push(...(t.tokens as Tokens.Generic[]));
    } else {
      out.push(t);
    }
  }
  return out;
}

export function renderMarkdown(md: string, builder: ContentBuilder): void {
  const tokens = marked.lexer(md);
  for (const token of tokens) {
    renderToken(token, builder);
  }
}
