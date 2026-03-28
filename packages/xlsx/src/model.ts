import type { PackageGraph } from '@ooxml/core';

export interface WorkbookSheet {
  name: string;
  uri: string;
  rows: WorksheetRow[];
}

export interface WorksheetRow {
  index: number;
  cells: WorksheetCell[];
}

export interface WorksheetCell {
  reference: string;
  type: string;
  value: string;
  formula?: string;
  styleIndex?: number;
}

export interface XlsxWorkbook {
  kind: 'xlsx';
  packageGraph: PackageGraph;
  sheets: WorkbookSheet[];
  sharedStrings: string[];
}
