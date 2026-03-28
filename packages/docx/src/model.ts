import type { PackageGraph } from '@ooxml/core';

export interface DocxRun {
  text: string;
  bold: boolean;
  italic: boolean;
}

export interface DocxParagraph {
  text: string;
  styleId?: string;
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

export interface DocxDocument {
  kind: 'docx';
  packageGraph: PackageGraph;
  stories: DocxStory[];
  comments: DocxComment[];
}
