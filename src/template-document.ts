import PizZip from "pizzip";
import * as fs from "fs";
import type { StylesConfig } from "./styles";
import { ContentBuilder } from "./content-builder";
import { createNumberingXml } from "./xml";

export interface TemplateDocumentConfig {
  templatePath: string;
  styles: StylesConfig;
  title?: string;
  date?: string;
  author?: string;
}

export class TemplateDocument {
  private zip: PizZip;
  private config: TemplateDocumentConfig;
  private contentBuilder: ContentBuilder;

  constructor(config: TemplateDocumentConfig) {
    this.config = config;
    this.contentBuilder = new ContentBuilder(config.styles);
    const templateContent = fs.readFileSync(config.templatePath, "binary");
    this.zip = new PizZip(templateContent);
  }

  get content(): ContentBuilder {
    return this.contentBuilder;
  }

  addTitlePage(): this {
    const c = this.contentBuilder;
    c.spacer().spacer().spacer()
      .title(this.config.title || "Document")
      .subtitle(this.config.date || "")
      .spacer()
      .spacer();
    if (this.config.author) {
      c.labeledPara("Prepared by: ", this.config.author);
    }
    c.pageBreak();
    return this;
  }

  save(outputPath: string): void {
    const docXml = this.zip.file("word/document.xml")?.asText();
    if (!docXml) throw new Error("Could not find document.xml in template");

    const bodyMatch = docXml.match(/(<w:body[^>]*>)([\s\S]*?)(<\/w:body>)/);
    if (!bodyMatch) throw new Error("Could not find body in document.xml");

    const [fullMatch, bodyOpen, existingContent, bodyClose] = bodyMatch;
    const sectPrMatch = existingContent.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
    const sectPr = sectPrMatch ? sectPrMatch[0] : "";

    const newContent = this.contentBuilder.toXml();
    const newBody = `${bodyOpen}\n${newContent}\n${sectPr}\n${bodyClose}`;
    const newDocXml = docXml.replace(fullMatch, newBody);

    this.zip.file("word/document.xml", newDocXml);
    this.zip.file("word/numbering.xml", createNumberingXml(this.config.styles));

    let contentTypes = this.zip.file("[Content_Types].xml")?.asText() || "";
    if (!contentTypes.includes("/word/numbering.xml")) {
      contentTypes = contentTypes.replace(
        "</Types>",
        '  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n</Types>'
      );
      this.zip.file("[Content_Types].xml", contentTypes);
    }

    let docRels = this.zip.file("word/_rels/document.xml.rels")?.asText() || "";
    if (!docRels.includes("numbering.xml")) {
      const rIdMatches = docRels.matchAll(/rId(\d+)/g);
      let maxId = 0;
      for (const match of rIdMatches) {
        maxId = Math.max(maxId, parseInt(match[1], 10));
      }
      const newRId = `rId${maxId + 1}`;
      docRels = docRels.replace(
        "</Relationships>",
        `  <Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>\n</Relationships>`
      );
      this.zip.file("word/_rels/document.xml.rels", docRels);
    }

    const output = this.zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
    fs.writeFileSync(outputPath, output);
  }
}
