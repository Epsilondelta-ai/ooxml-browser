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

export interface DocxRevision {
  kind: 'insertion' | 'deletion';
  id?: string;
  author?: string;
  date?: string;
  text: string;
}

export interface DocxParagraph {
  text: string;
  styleId?: string;
  numbering?: DocxParagraphNumbering;
  revisions: DocxRevision[];
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

export interface DocxMedia {
  relationshipId: string;
  targetUri: string;
  type: 'image' | 'embeddedObject';
  name?: string;
  progId?: string;
}

export type DocxBlock =
  | { kind: 'paragraph'; paragraph: DocxParagraph }
  | { kind: 'table'; table: DocxTable };

export interface DocxStory {
  kind: 'document' | 'header' | 'footer';
  uri: string;
  blocks: DocxBlock[];
  paragraphs: DocxParagraph[];
  tables: DocxTable[];
  media: DocxMedia[];
}

export interface DocxSectionPageSize {
  width: number;
  height: number;
}

export interface DocxSectionPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface DocxHeaderFooterReference {
  type: 'default' | 'first' | 'even';
  relationshipId: string;
  targetUri?: string;
}

export interface DocxSection {
  source: 'body' | 'paragraph';
  pageSize?: DocxSectionPageSize;
  pageMargins?: DocxSectionPageMargins;
  headerReferences: DocxHeaderFooterReference[];
  footerReferences: DocxHeaderFooterReference[];
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
  sections: DocxSection[];
}
