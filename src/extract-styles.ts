import PizZip from "pizzip";
import * as fs from "fs";
import { defaultStyles, type StylesConfig } from "./styles";

interface RawStyle {
  id: string;
  name: string;
  basedOn?: string;
  font?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  spacingBefore?: number;
  spacingAfter?: number;
}

interface ResolvedStyle {
  id: string;
  name: string;
  font?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  spacingBefore?: number;
  spacingAfter?: number;
}

interface DocDefaults {
  font?: string;
  size?: number;
}

function parseDocDefaults(stylesXml: string): DocDefaults {
  const m = stylesXml.match(/<w:docDefaults>([\s\S]*?)<\/w:docDefaults>/);
  if (!m) return {};
  const body = m[1];
  const font = body.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/)?.[1];
  const sizeMatch = body.match(/<w:sz\s+w:val="(\d+)"/);
  return { font, size: sizeMatch ? parseInt(sizeMatch[1], 10) : undefined };
}

function parseRawStyles(stylesXml: string): Map<string, RawStyle> {
  const map = new Map<string, RawStyle>();
  const styleMatches = stylesXml.matchAll(
    /<w:style\s+w:type="paragraph"[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g
  );
  for (const match of styleMatches) {
    const id = match[1];
    const body = match[2];
    const s: RawStyle = { id, name: id };

    const nameMatch = body.match(/<w:name\s+w:val="([^"]+)"/);
    if (nameMatch) s.name = nameMatch[1];

    const basedOnMatch = body.match(/<w:basedOn\s+w:val="([^"]+)"/);
    if (basedOnMatch) s.basedOn = basedOnMatch[1];

    const fontMatch = body.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/);
    if (fontMatch) s.font = fontMatch[1];

    const sizeMatch = body.match(/<w:sz\s+w:val="(\d+)"/);
    if (sizeMatch) s.size = parseInt(sizeMatch[1], 10);

    const colorMatch = body.match(/<w:color\s+w:val="([^"]+)"/);
    if (colorMatch && colorMatch[1] !== "auto") s.color = colorMatch[1];

    if (body.includes("<w:b/>") || body.includes('<w:b w:val="true"')) s.bold = true;
    if (body.includes("<w:i/>") || body.includes('<w:i w:val="true"')) s.italic = true;

    const sbMatch = body.match(/<w:spacing[^>]*w:before="(\d+)"/);
    if (sbMatch) s.spacingBefore = parseInt(sbMatch[1], 10);

    const saMatch = body.match(/<w:spacing[^>]*w:after="(\d+)"/);
    if (saMatch) s.spacingAfter = parseInt(saMatch[1], 10);

    map.set(id, s);
  }
  return map;
}

function resolveStyle(
  id: string,
  rawStyles: Map<string, RawStyle>,
  docDefaults: DocDefaults,
  seen: Set<string> = new Set()
): ResolvedStyle | undefined {
  if (seen.has(id)) return undefined;
  seen.add(id);
  const raw = rawStyles.get(id);
  if (!raw) return undefined;
  const parent = raw.basedOn
    ? resolveStyle(raw.basedOn, rawStyles, docDefaults, seen)
    : undefined;
  return {
    id: raw.id,
    name: raw.name,
    font: raw.font ?? parent?.font ?? docDefaults.font,
    size: raw.size ?? parent?.size ?? docDefaults.size,
    color: raw.color ?? parent?.color,
    bold: raw.bold ?? parent?.bold,
    italic: raw.italic ?? parent?.italic,
    spacingBefore: raw.spacingBefore ?? parent?.spacingBefore,
    spacingAfter: raw.spacingAfter ?? parent?.spacingAfter,
  };
}

function findStyleId(
  rawStyles: Map<string, RawStyle>,
  id: string,
  altName: string
): string | undefined {
  if (rawStyles.has(id)) return id;
  const altLower = altName.toLowerCase();
  for (const [sid, raw] of rawStyles) {
    if (raw.name.toLowerCase() === altLower) return sid;
  }
  return undefined;
}

const SYS_COLOR_FALLBACK: Record<string, string> = {
  windowText: "000000",
  window: "FFFFFF",
};

function parseThemeColors(themeXml: string): Record<string, string> {
  const themeColors: Record<string, string> = {};
  const schemeMatch = themeXml.match(/<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/);
  if (!schemeMatch) return themeColors;
  const schemeContent = schemeMatch[1];
  const tags = [
    "dk1", "dk2", "lt1", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
  ];
  for (const tag of tags) {
    const elMatch = schemeContent.match(
      new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`)
    );
    if (!elMatch) continue;
    const inner = elMatch[1];

    const srgb = inner.match(/<a:srgbClr\s+val="([^"]+)"/);
    if (srgb) {
      themeColors[tag] = srgb[1];
      continue;
    }
    const sysLast = inner.match(/<a:sysClr[^>]*lastClr="([^"]+)"/);
    if (sysLast) {
      themeColors[tag] = sysLast[1];
      continue;
    }
    const sysVal = inner.match(/<a:sysClr[^>]*val="([^"]+)"/);
    if (sysVal && SYS_COLOR_FALLBACK[sysVal[1]]) {
      themeColors[tag] = SYS_COLOR_FALLBACK[sysVal[1]];
    }
  }
  return themeColors;
}

export function extractStyles(templatePath: string): StylesConfig {
  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);
  const styles: StylesConfig = JSON.parse(JSON.stringify(defaultStyles));

  const themeXml = zip.file("word/theme/theme1.xml")?.asText();
  const themeColors = themeXml ? parseThemeColors(themeXml) : {};

  const stylesXml = zip.file("word/styles.xml")?.asText();
  if (!stylesXml) return styles;

  const docDefaults = parseDocDefaults(stylesXml);
  const rawStyles = parseRawStyles(stylesXml);

  const resolve = (id: string, altName: string): ResolvedStyle | undefined => {
    const sid = findStyleId(rawStyles, id, altName);
    return sid ? resolveStyle(sid, rawStyles, docDefaults) : undefined;
  };

  const title = resolve("Title", "Title");
  const h1 = resolve("Heading1", "Heading 1");
  const h2 = resolve("Heading2", "Heading 2");
  const h3 = resolve("Heading3", "Heading 3");
  const normal = resolve("Normal", "Normal");

  if (themeColors.accent1) styles.colors.primary = themeColors.accent1;
  if (themeColors.accent2) styles.colors.secondary = themeColors.accent2;
  if (themeColors.accent3) styles.colors.accent = themeColors.accent3;
  if (normal?.color) styles.colors.text = normal.color;
  else if (themeColors.dk1) styles.colors.text = themeColors.dk1;
  if (themeColors.lt1) styles.colors.background = themeColors.lt1;

  if (h1?.color) styles.colors.heading1 = h1.color;
  else if (themeColors.accent1) styles.colors.heading1 = themeColors.accent1;
  if (h2?.color) styles.colors.heading2 = h2.color;
  else if (themeColors.accent2) styles.colors.heading2 = themeColors.accent2;
  if (h3?.color) styles.colors.heading3 = h3.color;

  const headingFont = h1?.font || title?.font || normal?.font || docDefaults.font;
  const bodyFont = normal?.font || docDefaults.font;
  if (headingFont) styles.fonts.heading = headingFont;
  if (bodyFont) styles.fonts.body = bodyFont;

  if (title?.size) styles.fontSizes.title = title.size;
  if (h1?.size) styles.fontSizes.h1 = h1.size;
  if (h2?.size) styles.fontSizes.h2 = h2.size;
  if (h3?.size) styles.fontSizes.h3 = h3.size;

  const bodySize = normal?.size ?? docDefaults.size;
  if (bodySize) {
    styles.fontSizes.body = bodySize;
    styles.fontSizes.label = bodySize;
    styles.fontSizes.small = Math.round(bodySize * 0.9);
    styles.fontSizes.caption = Math.round(bodySize * 0.8);
  }

  if (normal?.spacingAfter !== undefined) styles.spacing.paragraph.after = normal.spacingAfter;
  if (h1?.spacingBefore !== undefined) styles.spacing.heading.h1.before = h1.spacingBefore;
  if (h1?.spacingAfter !== undefined) styles.spacing.heading.h1.after = h1.spacingAfter;
  if (h2?.spacingBefore !== undefined) styles.spacing.heading.h2.before = h2.spacingBefore;
  if (h2?.spacingAfter !== undefined) styles.spacing.heading.h2.after = h2.spacingAfter;
  if (h3?.spacingBefore !== undefined) styles.spacing.heading.h3.before = h3.spacingBefore;
  if (h3?.spacingAfter !== undefined) styles.spacing.heading.h3.after = h3.spacingAfter;

  styles.listStyles.bullet.font = styles.fonts.body;
  styles.listStyles.bullet.size = styles.fontSizes.body;
  styles.listStyles.bullet.color = styles.colors.text;
  styles.listStyles.numbered.font = styles.fonts.body;
  styles.listStyles.numbered.size = styles.fontSizes.body;
  styles.listStyles.numbered.color = styles.colors.text;

  const td = styles.tableStyles.default;
  td.headerBackground = styles.colors.primary;
  td.headerFont = styles.fonts.body;
  td.headerSize = styles.fontSizes.body;
  td.bodyFont = styles.fonts.body;
  td.bodySize = styles.fontSizes.body;
  td.bodyTextColor = styles.colors.text;
  td.rowOdd = styles.colors.backgroundAlt;

  const ts = styles.tableStyles.subtle;
  ts.headerFont = styles.fonts.body;
  ts.headerSize = styles.fontSizes.body;
  ts.bodyFont = styles.fonts.body;
  ts.bodySize = styles.fontSizes.body;
  ts.bodyTextColor = styles.colors.text;
  ts.headerTextColor = styles.colors.text;

  return styles;
}
