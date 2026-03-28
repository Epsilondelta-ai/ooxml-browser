import type { PackageGraph } from '@ooxml/core';

export interface DocxRun {
  text: string;
  bold: boolean;
  italic: boolean;
}

export interface DocxParagraphNumbering {
  numId: string;
  level: number;
}

export interface DocxParagraph {
  text: string;
  styleId?: string;
  numbering?: DocxParagraphNumbering;
  runs: DocxRun[];
}

export interface DocxTableCell {
  text: string;
}

export interface DocxTableRow {
  cells: DocxTableCell[];
}

export interface DocxTable {
  rows: DocxTableRow[];
}

export interface DocxComment {
  id: string;
  author?: string;
  text: string;
}

export interface DocxStory {
  kind: 'document' | 'header' | 'footer';
  uri: string;
  paragraphs: DocxParagraph[];
  tables: DocxTable[];
}

export interface DocxNumberingLevel {
  level: number;
  format: string;
  text: string;
  start: number;
}

export interface DocxAbstractNumbering {
  id: string;
  levels: Record<number, DocxNumberingLevel>;
}

export interface DocxNumberingInstance {
  id: string;
  abstractNumId: string;
}

export interface DocxNumbering {
  partUri?: string;
  abstractNums: Record<string, DocxAbstractNumbering>;
  nums: Record<string, DocxNumberingInstance>;
}

export interface DocxStyle {
  id: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering' | 'unknown';
  name?: string;
  basedOn?: string;
  isDefault: boolean;
  bold?: boolean;
  italic?: boolean;
}

export interface DocxDocument {
  kind: 'docx';
  packageGraph: PackageGraph;
  stories: DocxStory[];
  comments: DocxComment[];
  styles: Record<string, DocxStyle>;
  numbering: DocxNumbering;
}
