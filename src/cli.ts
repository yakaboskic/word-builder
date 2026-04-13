import * as fs from "fs";
import * as path from "path";
import { extractStyles } from "./extract-styles";
import { loadStyles, saveStyles, type StylesConfig } from "./styles";
import { TemplateDocument } from "./template-document";
import { ContentBuilder } from "./content-builder";
import { renderMarkdown, parseFrontmatter } from "./markdown";

function printHelp(): void {
  console.log(`word-builder — generate Word documents from a template + content

Usage:
  word-builder extract <template.docx> [-o <styles.json>]
  word-builder build   <template.docx> <content.(md|ts|js|mjs)> [options]

Extract options:
  -o, --output <path>     Output JSON path (default: <template>.styles.json)

Build options:
  -s, --styles <path>     Styles JSON (default: auto-extract from template)
  -o, --output <path>     Output docx (default: <content>.out.docx)
  --title <text>          Document title (adds title page if set)
  --date <text>           Document date
  --author <text>         Document author
  --title-page            Force a title page
  --no-title-page         Suppress title page

Content files:
  *.md / *.markdown       Parsed as Markdown. Frontmatter title/date/author is honored.
  *.ts / *.js / *.mjs     Must export a default function: (builder: ContentBuilder) => void.
                          Optional 'meta' export: { title, date, author }.

Examples:
  word-builder extract template.docx -o styles.json
  word-builder build template.docx notes.md -o notes.docx
  word-builder build template.docx chapter.ts -s styles.json -o chapter.docx
`);
}

interface ParsedArgs {
  positional: string[];
  flags: Record<string, string | boolean>;
}

function parseArgs(argv: string[]): ParsedArgs {
  const positional: string[] = [];
  const flags: Record<string, string | boolean> = {};
  const shortMap: Record<string, string> = { o: "output", s: "styles" };

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith("--")) {
      let key = a.slice(2);
      if (key.startsWith("no-")) {
        flags[key.slice(3)] = false;
        continue;
      }
      const next = argv[i + 1];
      if (next !== undefined && !next.startsWith("-")) {
        flags[key] = next;
        i++;
      } else {
        flags[key] = true;
      }
    } else if (a.startsWith("-") && a.length === 2) {
      const key = shortMap[a.slice(1)] ?? a.slice(1);
      const next = argv[i + 1];
      if (next !== undefined && !next.startsWith("-")) {
        flags[key] = next;
        i++;
      } else {
        flags[key] = true;
      }
    } else {
      positional.push(a);
    }
  }
  return { positional, flags };
}

function runExtract(args: ParsedArgs): void {
  const template = args.positional[0];
  if (!template) {
    console.error("extract: missing <template.docx>");
    process.exit(1);
  }
  if (!fs.existsSync(template)) {
    console.error(`Template not found: ${template}`);
    process.exit(1);
  }
  const output =
    (args.flags.output as string) ||
    `${path.basename(template, path.extname(template))}.styles.json`;

  const styles = extractStyles(template);
  saveStyles(output, styles);

  console.log(`Extracted styles → ${output}`);
  console.log(`  fonts:   heading=${styles.fonts.heading}, body=${styles.fonts.body}`);
  console.log(
    `  sizes:   title=${styles.fontSizes.title / 2}pt  h1=${styles.fontSizes.h1 / 2}pt  h2=${styles.fontSizes.h2 / 2}pt  body=${styles.fontSizes.body / 2}pt`
  );
  console.log(`  primary: #${styles.colors.primary}`);
}

async function runBuild(args: ParsedArgs): Promise<void> {
  const template = args.positional[0];
  const contentPath = args.positional[1];
  if (!template || !contentPath) {
    console.error("build: usage — word-builder build <template.docx> <content>");
    process.exit(1);
  }
  if (!fs.existsSync(template)) {
    console.error(`Template not found: ${template}`);
    process.exit(1);
  }
  if (!fs.existsSync(contentPath)) {
    console.error(`Content not found: ${contentPath}`);
    process.exit(1);
  }

  let styles: StylesConfig;
  if (args.flags.styles) {
    styles = loadStyles(args.flags.styles as string);
  } else {
    styles = extractStyles(template);
  }

  const output =
    (args.flags.output as string) ||
    `${path.basename(contentPath, path.extname(contentPath))}.out.docx`;

  const ext = path.extname(contentPath).toLowerCase();

  let meta: { title?: string; date?: string; author?: string } = {};
  let buildFn: ((builder: ContentBuilder) => void | Promise<void>) | null = null;

  if (ext === ".md" || ext === ".markdown") {
    const md = fs.readFileSync(contentPath, "utf-8");
    const parsed = parseFrontmatter(md);
    meta = parsed.meta;
    buildFn = (builder) => renderMarkdown(parsed.body, builder);
  } else if (ext === ".ts" || ext === ".js" || ext === ".mjs" || ext === ".cjs") {
    const mod = await import(path.resolve(contentPath));
    if (mod.meta) meta = mod.meta;
    if (typeof mod.default === "function") buildFn = mod.default;
    else if (typeof mod.build === "function") buildFn = mod.build;
    else {
      console.error(`${contentPath} must export a default function (builder: ContentBuilder) => void`);
      process.exit(1);
    }
  } else {
    console.error(`Unsupported content extension: ${ext}`);
    process.exit(1);
  }

  const title = (args.flags.title as string) ?? meta.title;
  const date = (args.flags.date as string) ?? meta.date;
  const author = (args.flags.author as string) ?? meta.author;

  const doc = new TemplateDocument({ templatePath: template, styles, title, date, author });

  const forced = args.flags["title-page"];
  const suppressed = forced === false;
  const shouldAddTitlePage = forced === true || (!suppressed && !!title);
  if (shouldAddTitlePage) doc.addTitlePage();

  if (buildFn) await buildFn(doc.content);

  doc.save(output);
  console.log(`Built → ${output}`);
}

async function main(): Promise<void> {
  const [cmd, ...rest] = process.argv.slice(2);
  const parsed = parseArgs(rest);
  switch (cmd) {
    case "extract":
      runExtract(parsed);
      break;
    case "build":
      await runBuild(parsed);
      break;
    case undefined:
    case "help":
    case "--help":
    case "-h":
      printHelp();
      break;
    default:
      console.error(`Unknown command: ${cmd}`);
      printHelp();
      process.exit(1);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
