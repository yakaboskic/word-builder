import * as fs from "fs";

export interface ListStyle {
  font: string;
  size: number;
  color: string;
}

export interface TableStyle {
  headerBackground: string;
  headerTextColor: string;
  headerFont: string;
  headerSize: number;
  headerBold: boolean;
  bodyTextColor: string;
  bodyFont: string;
  bodySize: number;
  rowEven: string;
  rowOdd: string;
  border: string;
  borderWidth: number;
}

export interface NoteStyle {
  background: string;
  borderColor: string;
}

export interface StylesConfig {
  colors: {
    primary: string;
    secondary: string;
    accent: string;
    text: string;
    textLight: string;
    textMuted: string;
    background: string;
    backgroundAlt: string;
    backgroundMuted: string;
    border: string;
    borderLight: string;
    success: string;
    warning: string;
    warningBg: string;
    error: string;
    info: string;
    heading1: string;
    heading2: string;
    heading3: string;
  };
  fonts: {
    heading: string;
    body: string;
    mono: string;
  };
  fontSizes: {
    title: number;
    h1: number;
    h2: number;
    h3: number;
    body: number;
    small: number;
    caption: number;
    label: number;
  };
  spacing: {
    paragraph: { after: number; afterSmall: number; afterLarge: number };
    heading: {
      h1: { before: number; after: number };
      h2: { before: number; after: number };
      h3: { before: number; after: number };
    };
    lineHeight: number;
    lineHeightTight: number;
  };
  listStyles: {
    bullet: ListStyle;
    numbered: ListStyle;
  };
  tableStyles: {
    default: TableStyle;
    subtle: TableStyle;
  };
  noteStyles: {
    warning: NoteStyle;
    info: NoteStyle;
    success: NoteStyle;
    error: NoteStyle;
  };
}

export type TableStyleName = keyof StylesConfig["tableStyles"];
export type NoteStyleName = keyof StylesConfig["noteStyles"];

export function loadStyles(path: string): StylesConfig {
  return JSON.parse(fs.readFileSync(path, "utf-8")) as StylesConfig;
}

export function saveStyles(path: string, styles: StylesConfig): void {
  fs.writeFileSync(path, JSON.stringify(styles, null, 2));
}

export const defaultStyles: StylesConfig = {
  colors: {
    primary: "1F4E79",
    secondary: "2E75B6",
    accent: "5B9BD5",
    text: "000000",
    textLight: "666666",
    textMuted: "999999",
    background: "FFFFFF",
    backgroundAlt: "F2F2F2",
    backgroundMuted: "F5F5F5",
    border: "D9D9D9",
    borderLight: "E5E5E5",
    success: "28A745",
    warning: "FFC107",
    warningBg: "FFF3CD",
    error: "DC3545",
    info: "17A2B8",
    heading1: "1F4E79",
    heading2: "2E75B6",
    heading3: "000000",
  },
  fonts: {
    heading: "Calibri Light",
    body: "Calibri",
    mono: "Consolas",
  },
  fontSizes: {
    title: 52,
    h1: 32,
    h2: 26,
    h3: 24,
    body: 22,
    small: 20,
    caption: 18,
    label: 22,
  },
  spacing: {
    paragraph: { after: 200, afterSmall: 120, afterLarge: 400 },
    heading: {
      h1: { before: 400, after: 200 },
      h2: { before: 300, after: 150 },
      h3: { before: 200, after: 100 },
    },
    lineHeight: 276,
    lineHeightTight: 240,
  },
  listStyles: {
    bullet: { font: "Calibri", size: 22, color: "000000" },
    numbered: { font: "Calibri", size: 22, color: "000000" },
  },
  tableStyles: {
    default: {
      headerBackground: "1F4E79",
      headerTextColor: "FFFFFF",
      headerFont: "Calibri",
      headerSize: 22,
      headerBold: true,
      bodyTextColor: "000000",
      bodyFont: "Calibri",
      bodySize: 22,
      rowEven: "FFFFFF",
      rowOdd: "F2F2F2",
      border: "D9D9D9",
      borderWidth: 4,
    },
    subtle: {
      headerBackground: "F2F2F2",
      headerTextColor: "000000",
      headerFont: "Calibri",
      headerSize: 22,
      headerBold: true,
      bodyTextColor: "000000",
      bodyFont: "Calibri",
      bodySize: 22,
      rowEven: "FFFFFF",
      rowOdd: "FFFFFF",
      border: "E5E5E5",
      borderWidth: 4,
    },
  },
  noteStyles: {
    warning: { background: "FFF3CD", borderColor: "FFC107" },
    info: { background: "E7F3FE", borderColor: "17A2B8" },
    success: { background: "D4EDDA", borderColor: "28A745" },
    error: { background: "F8D7DA", borderColor: "DC3545" },
  },
};
