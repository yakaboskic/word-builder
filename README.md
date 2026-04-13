# word-builder

Generate Word documents from any `.docx` template plus content written in Markdown, TypeScript, or JavaScript. The template contributes headers, footers, theme, page layout, and typography; you contribute the body.

## How it works

A `.docx` is a zip of XML parts. `word-builder` does two things against that zip:

1. **Extract styles** — reads `word/styles.xml` and `word/theme/theme1.xml`, resolves the `basedOn` chain and `docDefaults` for `Title` / `Heading1-3` / `Normal`, and writes a plain JSON file describing fonts, sizes, colors, spacing, and list/table styling.
2. **Build document** — loads the template, preserves its `<w:sectPr>` (so headers, footers, margins, theme, and the numbering reference in `[Content_Types].xml` stay intact), replaces the body with content generated from the fluent `ContentBuilder` API, and writes a new `.docx`.

Content can be authored in whichever format fits the task:

- **Markdown** — for simple prose, lists, and tables. Frontmatter `title` / `date` / `author` drives the optional cover page.
- **TypeScript / JavaScript** — when you need the full `ContentBuilder` API (per-element style overrides, tables, notes, image placeholders, custom spacing).

## Install

```bash
pnpm install
```

## CLI

```bash
# Extract styles from a template into a JSON file
pnpm exec tsx src/cli.ts extract <template.docx> [-o styles.json]

# Build a document
pnpm exec tsx src/cli.ts build <template.docx> <content.(md|ts|js)> [options]
```

Shortcuts are defined in `package.json`:

```bash
pnpm extract <template.docx> -o styles.json
pnpm build   <template.docx> content.md -o out.docx
```

### Build options

| Flag | Description |
|---|---|
| `-s, --styles <path>` | Styles JSON to apply. If omitted, styles are auto-extracted from the template. |
| `-o, --output <path>` | Output `.docx` path. Defaults to `<content>.out.docx`. |
| `--title <text>` | Document title (adds a title page when set). |
| `--date <text>` | Document date (shown on title page). |
| `--author <text>` | Document author (shown on title page). |
| `--title-page` | Force a title page even without a title. |
| `--no-title-page` | Suppress the title page. |

CLI flags override whatever came from Markdown frontmatter or a TS module's `meta` export.

## Markdown content

Frontmatter is optional. Supported block elements: headings (`#`, `##`, `###`), paragraphs, ordered and unordered lists, tables, blockquotes (rendered as inline "Note:" text), fenced code blocks (rendered in a mono font), and `---` (rendered as a page break). Inline `**bold**`, `*italic*`, and `` `code` `` are preserved.

```markdown
---
title: Quarterly Report
date: April 2026
author: Chase
---

# 1. Introduction

This paragraph supports **bold**, *italic*, and `code` inline.

## 1.1 Bullets

- First point
- Second point with a **bold** word

## 1.2 Table

| Column A | Column B |
| --- | --- |
| 1 | 2 |
```

## TypeScript / JavaScript content

Export a default function taking a `ContentBuilder`. Optionally export a `meta` object for title-page fields.

```ts
// content.ts
import type { ContentBuilder } from "./src/content-builder";

export const meta = {
  title: "Quarterly Report",
  date: "April 2026",
  author: "Chase",
};

export default function build(c: ContentBuilder): void {
  c.h1("1. Introduction")
    .p("Opening paragraph.")
    .h2("1.1 Key metrics")
    .table(
      ["Metric", "Value"],
      [
        ["Requests", "1,240"],
        ["Errors", "3"],
      ]
    )
    .h2("1.2 Details")
    .bullet("First detail", "Key: ")
    .bullet("Second detail")
    .p("Closing paragraph.", { italic: true, color: "666666" });
}
```

## ContentBuilder API

All methods return `this` for chaining. The `override` parameter (last argument on text methods) accepts `{ font, size, color, bold, italic, alignment, spacingBefore, spacingAfter }`.

| Method | Description |
|---|---|
| `title(text, override?)` | Centered title |
| `subtitle(text, override?)` | Centered subtitle |
| `h1(text, override?)` / `h2` / `h3` | Headings — apply template's heading style, suppress its auto-numbering |
| `p(text \| RichText[], override?)` | Paragraph. Accepts rich text parts for mixed formatting. |
| `labeledPara(label, text, override?)` | Bold-label prefix + body text |
| `label(text, override?)` | Inline bold label |
| `bullet(text, boldPrefix?, override?)` | Bulleted list item |
| `numbered(text, boldPrefix?, override?)` | Numbered list item |
| `note(text, type?, label?, override?)` | Inline "Note:" callout |
| `imagePlaceholder(caption, description, override?)` | Placeholder block for a figure |
| `table(headers, rows, styleName?)` | Table — `"default"` or `"subtle"` |
| `spacer()` | Empty paragraph |
| `pageBreak()` | Page break |

Sizes are in half-points (e.g. `size: 22` → 11pt). Colors are hex strings without `#`. Spacing values are in twentieths of a point (e.g. `spacingAfter: 200` → 10pt).

## Style JSON

Extraction produces a flat JSON file with these top-level keys:

- `colors` — primary, secondary, accent, text, background, heading1-3, plus semantic colors (success/warning/error/info)
- `fonts` — `heading`, `body`, `mono`
- `fontSizes` — `title`, `h1`, `h2`, `h3`, `body`, `small`, `caption`, `label` (half-points)
- `spacing` — paragraph and per-heading before/after values
- `listStyles` — bullet and numbered list formatting
- `tableStyles` — `default` and `subtle` variants
- `noteStyles` — warning/info/success/error palettes

Edit the JSON to customize anything; the next `build` will pick up the changes without re-extracting.

## Project layout

```
src/
├── cli.ts                # CLI entry (extract, build)
├── extract-styles.ts     # Parse styles.xml + theme1.xml → StylesConfig
├── styles.ts             # StylesConfig type, defaults, load/save
├── content-builder.ts    # Fluent API for building body content
├── template-document.ts  # Loads template, replaces body, saves
├── markdown.ts           # Markdown + frontmatter → ContentBuilder calls
└── xml.ts                # OOXML helpers (paragraphs, runs, tables, numbering)
```

## Caveats

- **Body is replaced, not merged.** Building against a template discards whatever content was in the template's `<w:body>`; only `<w:sectPr>` (and everything it references: headers, footers, theme, images) is preserved. Templates you use should be "shells" without body content you want to keep.
- **`numbering.xml` is rewritten.** The generated numbering defines only `numId=1` (bullets) and `numId=2` (numbers). If your template's headers or footers reference other `numId` values, they may lose their list formatting.
- **Heading auto-numbering is suppressed.** When the template's heading styles have built-in list numbering, it's turned off at the paragraph level so your content's manual numbering isn't duplicated.
- **Blockquotes are not colored callout boxes.** They render as italic "Note:" text, not as shaded paragraphs with borders — shaded backgrounds aren't wired up yet.
