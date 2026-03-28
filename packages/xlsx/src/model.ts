import type { PackageGraph } from '@ooxml/core';

export interface XlsxFrozenPane {
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  state?: string;
}

export interface XlsxSelection {
  activeCell?: string;
  sqref?: string;
}

export interface XlsxPageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export interface XlsxPageSetup {
  orientation?: string;
  paperSize?: number;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
}

export interface XlsxDefinedName {
  name: string;
  reference: string;
  scopeSheetId?: number;
}

export interface XlsxTable {
  name: string;
  range: string;
  partUri: string;
}

export interface XlsxComment {
  reference: string;
  author?: string;
  text: string;
}

export interface WorkbookSheet {
  name: string;
  uri: string;
  sheetId: number;
  relationshipId: string;
  rows: WorksheetRow[];
  mergedRanges: string[];
  frozenPane?: XlsxFrozenPane;
  selection?: XlsxSelection;
  pageMargins?: XlsxPageMargins;
  pageSetup?: XlsxPageSetup;
  tables: XlsxTable[];
  comments: XlsxComment[];
}

export interface WorksheetRow {
  index: number;
  cells: WorksheetCell[];
}

export interface XlsxNumberFormat {
  id: number;
  code: string;
}

export interface XlsxCellFormat {
  id: number;
  numFmtId: number;
}

export interface XlsxStyleTable {
  partUri?: string;
  numberFormats: Record<number, XlsxNumberFormat>;
  cellFormats: Record<number, XlsxCellFormat>;
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
  styles: XlsxStyleTable;
  definedNames: XlsxDefinedName[];
}
