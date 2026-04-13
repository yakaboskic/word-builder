import type { StylesConfig, TableStyleName, NoteStyleName } from "./styles";
import {
  createParagraph,
  createPageBreak,
  createTable,
  type RichText,
} from "./xml";

export interface TextStyleOverride {
  font?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  alignment?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
}

export class ContentBuilder {
  private items: string[] = [];

  constructor(private styles: StylesConfig) {}

  title(text: string, override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(text, {
        alignment: override?.alignment ?? "center",
        font: override?.font ?? s.fonts.heading,
        size: override?.size ?? s.fontSizes.title,
        color: override?.color ?? s.colors.primary,
        bold: override?.bold ?? true,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore,
        spacingAfter: override?.spacingAfter,
        suppressNumbering: true,
      })
    );
    return this;
  }

  subtitle(text: string, override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(text, {
        alignment: override?.alignment ?? "center",
        font: override?.font ?? s.fonts.body,
        size: override?.size ?? s.fontSizes.body + 6,
        color: override?.color ?? s.colors.textLight,
        bold: override?.bold,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore,
        spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.afterLarge,
      })
    );
    return this;
  }

  h1(text: string, override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(text, {
        style: "Heading1",
        alignment: override?.alignment,
        font: override?.font ?? s.fonts.heading,
        size: override?.size ?? s.fontSizes.h1,
        color: override?.color ?? s.colors.heading1,
        bold: override?.bold ?? true,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore ?? s.spacing.heading.h1.before,
        spacingAfter: override?.spacingAfter ?? s.spacing.heading.h1.after,
        suppressNumbering: true,
      })
    );
    return this;
  }

  h2(text: string, override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(text, {
        style: "Heading2",
        alignment: override?.alignment,
        font: override?.font ?? s.fonts.heading,
        size: override?.size ?? s.fontSizes.h2,
        color: override?.color ?? s.colors.heading2,
        bold: override?.bold ?? true,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore ?? s.spacing.heading.h2.before,
        spacingAfter: override?.spacingAfter ?? s.spacing.heading.h2.after,
        suppressNumbering: true,
      })
    );
    return this;
  }

  h3(text: string, override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(text, {
        style: "Heading3",
        alignment: override?.alignment,
        font: override?.font ?? s.fonts.heading,
        size: override?.size ?? s.fontSizes.h3,
        color: override?.color ?? s.colors.heading3,
        bold: override?.bold ?? true,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore ?? s.spacing.heading.h3.before,
        spacingAfter: override?.spacingAfter ?? s.spacing.heading.h3.after,
        suppressNumbering: true,
      })
    );
    return this;
  }

  p(content: string | RichText[], override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(content, {
        alignment: override?.alignment,
        font: override?.font ?? s.fonts.body,
        size: override?.size ?? s.fontSizes.body,
        color: override?.color ?? s.colors.text,
        bold: override?.bold,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore,
        spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.after,
      })
    );
    return this;
  }

  labeledPara(label: string, text: string, override?: TextStyleOverride): this {
    return this.p([{ text: label, bold: true }, { text }], override);
  }

  label(text: string, override?: TextStyleOverride): this {
    const s = this.styles;
    this.items.push(
      createParagraph(text, {
        alignment: override?.alignment,
        font: override?.font ?? s.fonts.body,
        size: override?.size ?? s.fontSizes.label,
        color: override?.color ?? s.colors.text,
        bold: override?.bold ?? true,
        italic: override?.italic,
        spacingBefore: override?.spacingBefore,
        spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.afterSmall,
      })
    );
    return this;
  }

  bullet(
    text: string | RichText[],
    boldPrefix?: string,
    override?: TextStyleOverride
  ): this {
    const s = this.styles;
    const base = {
      alignment: override?.alignment,
      font: override?.font ?? s.listStyles.bullet.font,
      size: override?.size ?? s.listStyles.bullet.size,
      color: override?.color ?? s.listStyles.bullet.color,
      italic: override?.italic,
      bullet: true,
      spacingBefore: override?.spacingBefore,
      spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.afterSmall,
    };
    if (boldPrefix && typeof text === "string") {
      this.items.push(
        createParagraph([{ text: boldPrefix, bold: true }, { text }], base)
      );
    } else {
      this.items.push(createParagraph(text, { ...base, bold: override?.bold }));
    }
    return this;
  }

  numbered(
    text: string | RichText[],
    boldPrefix?: string,
    override?: TextStyleOverride
  ): this {
    const s = this.styles;
    const base = {
      alignment: override?.alignment,
      font: override?.font ?? s.listStyles.numbered.font,
      size: override?.size ?? s.listStyles.numbered.size,
      color: override?.color ?? s.listStyles.numbered.color,
      italic: override?.italic,
      numbered: true,
      spacingBefore: override?.spacingBefore,
      spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.afterSmall,
    };
    if (boldPrefix && typeof text === "string") {
      this.items.push(
        createParagraph([{ text: boldPrefix, bold: true }, { text }], base)
      );
    } else {
      this.items.push(createParagraph(text, { ...base, bold: override?.bold }));
    }
    return this;
  }

  note(
    text: string,
    type: NoteStyleName = "warning",
    label: string = "Note: ",
    override?: TextStyleOverride
  ): this {
    const s = this.styles;
    this.items.push(
      createParagraph(
        [
          { text: label, bold: true, italic: true },
          { text, italic: override?.italic ?? true },
        ],
        {
          alignment: override?.alignment,
          font: override?.font ?? s.fonts.body,
          size: override?.size ?? s.fontSizes.body,
          color: override?.color ?? s.colors.text,
          spacingBefore: override?.spacingBefore,
          spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.after,
        }
      )
    );
    return this;
  }

  imagePlaceholder(
    caption: string,
    description: string,
    override?: TextStyleOverride
  ): this {
    const s = this.styles;
    this.items.push(
      createParagraph(`[${caption}]`, {
        alignment: override?.alignment ?? "center",
        font: override?.font ?? s.fonts.body,
        size: override?.size ?? s.fontSizes.body,
        color: override?.color ?? s.colors.textLight,
        bold: override?.bold,
        italic: override?.italic ?? true,
        spacingBefore: override?.spacingBefore,
        spacingAfter: s.spacing.paragraph.afterSmall,
      })
    );
    this.items.push(
      createParagraph(description, {
        alignment: "center",
        font: s.fonts.body,
        size: s.fontSizes.caption,
        color: s.colors.textMuted,
        italic: true,
        spacingAfter: override?.spacingAfter ?? s.spacing.paragraph.after,
      })
    );
    return this;
  }

  table(headers: string[], rows: string[][], styleName: TableStyleName = "default"): this {
    this.items.push(createTable(headers, rows, this.styles, styleName));
    this.items.push(createParagraph("", { spacingAfter: this.styles.spacing.paragraph.after }));
    return this;
  }

  spacer(): this {
    this.items.push(createParagraph("", { spacingAfter: this.styles.spacing.paragraph.after }));
    return this;
  }

  pageBreak(): this {
    this.items.push(createPageBreak());
    return this;
  }

  toXml(): string {
    return this.items.join("\n");
  }
}
