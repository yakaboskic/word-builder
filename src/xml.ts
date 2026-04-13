import type { StylesConfig, TableStyleName } from "./styles";

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export interface RunOptions {
  font?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}

export interface ParaOptions extends RunOptions {
  style?: string;
  alignment?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
  bullet?: boolean;
  numbered?: boolean;
  suppressNumbering?: boolean;
}

export interface RichText {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  font?: string;
  size?: number;
}

export function createRunProps(options: RunOptions): string {
  const parts: string[] = [];
  if (options.font) parts.push(`<w:rFonts w:ascii="${options.font}" w:hAnsi="${options.font}"/>`);
  if (options.size) parts.push(`<w:sz w:val="${options.size}"/><w:szCs w:val="${options.size}"/>`);
  if (options.color) parts.push(`<w:color w:val="${options.color}"/>`);
  if (options.bold) parts.push(`<w:b/>`);
  if (options.italic) parts.push(`<w:i/>`);
  return parts.length > 0 ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

export function createParaProps(options: ParaOptions): string {
  const parts: string[] = [];
  if (options.style) parts.push(`<w:pStyle w:val="${options.style}"/>`);
  if (options.alignment) parts.push(`<w:jc w:val="${options.alignment}"/>`);
  if (options.spacingBefore !== undefined || options.spacingAfter !== undefined) {
    const before = options.spacingBefore !== undefined ? ` w:before="${options.spacingBefore}"` : "";
    const after = options.spacingAfter !== undefined ? ` w:after="${options.spacingAfter}"` : "";
    parts.push(`<w:spacing${before}${after}/>`);
  }
  if (options.bullet) {
    parts.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`);
  } else if (options.numbered) {
    parts.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>`);
  } else if (options.suppressNumbering) {
    parts.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>`);
  }
  return parts.length > 0 ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

export function createRun(text: string, options: RunOptions = {}): string {
  const rPr = createRunProps(options);
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

export function createParagraph(
  content: string | RichText[],
  options: ParaOptions = {}
): string {
  const pPr = createParaProps(options);
  let runs: string;
  if (typeof content === "string") {
    runs = createRun(content, options);
  } else {
    runs = content
      .map((part) =>
        createRun(part.text, {
          font: part.font ?? options.font,
          size: part.size ?? options.size,
          color: part.color ?? options.color,
          bold: part.bold ?? options.bold,
          italic: part.italic ?? options.italic,
        })
      )
      .join("");
  }
  return `<w:p>${pPr}${runs}</w:p>`;
}

export function createPageBreak(): string {
  return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
}

export function createTable(
  headers: string[],
  rows: string[][],
  styles: StylesConfig,
  styleName: TableStyleName = "default"
): string {
  const style = styles.tableStyles[styleName];

  const headerRPr = [
    `<w:rFonts w:ascii="${style.headerFont}" w:hAnsi="${style.headerFont}"/>`,
    `<w:sz w:val="${style.headerSize}"/><w:szCs w:val="${style.headerSize}"/>`,
    `<w:color w:val="${style.headerTextColor}"/>`,
    style.headerBold ? `<w:b/>` : "",
  ].join("");

  const bodyRPr = [
    `<w:rFonts w:ascii="${style.bodyFont}" w:hAnsi="${style.bodyFont}"/>`,
    `<w:sz w:val="${style.bodySize}"/><w:szCs w:val="${style.bodySize}"/>`,
    `<w:color w:val="${style.bodyTextColor}"/>`,
  ].join("");

  const headerRow = `<w:tr>${headers
    .map(
      (h) =>
        `<w:tc><w:tcPr><w:shd w:val="clear" w:fill="${style.headerBackground}"/></w:tcPr>` +
        `<w:p><w:r><w:rPr>${headerRPr}</w:rPr>` +
        `<w:t>${escapeXml(h)}</w:t></w:r></w:p></w:tc>`
    )
    .join("")}</w:tr>`;

  const dataRows = rows
    .map(
      (row, idx) =>
        `<w:tr>${row
          .map(
            (cell) =>
              `<w:tc><w:tcPr><w:shd w:val="clear" w:fill="${idx % 2 === 0 ? style.rowEven : style.rowOdd}"/></w:tcPr>` +
              `<w:p><w:r><w:rPr>${bodyRPr}</w:rPr><w:t>${escapeXml(cell)}</w:t></w:r></w:p></w:tc>`
          )
          .join("")}</w:tr>`
    )
    .join("");

  const borderSz = style.borderWidth;

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="5000" w:type="pct"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="${borderSz}" w:color="${style.border}"/>
        <w:left w:val="single" w:sz="${borderSz}" w:color="${style.border}"/>
        <w:bottom w:val="single" w:sz="${borderSz}" w:color="${style.border}"/>
        <w:right w:val="single" w:sz="${borderSz}" w:color="${style.border}"/>
        <w:insideH w:val="single" w:sz="${borderSz}" w:color="${style.border}"/>
        <w:insideV w:val="single" w:sz="${borderSz}" w:color="${style.border}"/>
      </w:tblBorders>
    </w:tblPr>
    ${headerRow}
    ${dataRows}
  </w:tbl>`;
}

export function createNumberingXml(styles: StylesConfig): string {
  const bulletStyle = styles.listStyles.bullet;
  const numberedStyle = styles.listStyles.numbered;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
             xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8226;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="${bulletStyle.font}" w:hAnsi="${bulletStyle.font}" w:hint="default"/>
        <w:sz w:val="${bulletStyle.size}"/>
        <w:szCs w:val="${bulletStyle.size}"/>
        <w:color w:val="${bulletStyle.color}"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="${numberedStyle.font}" w:hAnsi="${numberedStyle.font}" w:hint="default"/>
        <w:sz w:val="${numberedStyle.size}"/>
        <w:szCs w:val="${numberedStyle.size}"/>
        <w:color w:val="${numberedStyle.color}"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
}
