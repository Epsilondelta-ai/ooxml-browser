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

export interface XlsxThreadedCommentPerson {
  id: string;
  displayName: string;
}

export interface XlsxThreadedComment {
  id: string;
  reference: string;
  personId: string;
  parentId?: string;
  text: string;
  author?: string;
}

export interface XlsxChartDataLabels {
  position?: string;
  separator?: string;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showLegendKey?: boolean;
  showLeaderLines?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
}

export interface XlsxChartSeries {
  name: string;
  invertIfNegative?: boolean;
  markerSymbol?: string;
  markerSize?: number;
  explosion?: number;
}

export interface XlsxChart {
  relationshipId: string;
  drawingUri: string;
  drawingNameOccurrence: number;
  targetUri: string;
  name?: string;
  chartType?: string;
  scatterStyle?: string;
  bubbleScale?: number;
  showNegativeBubbles?: boolean;
  sizeRepresents?: string;
  smooth?: boolean;
  grouping?: string;
  overlap?: number;
  varyColors?: boolean;
  gapWidth?: number;
  title?: string;
  firstSliceAngle?: number;
  holeSize?: number;
  plotVisibleOnly?: boolean;
  displayBlanksAs?: string;
  legendPosition?: string;
  categoryAxisTitle?: string;
  categoryAxisPosition?: string;
  valueAxisTitle?: string;
  valueAxisPosition?: string;
  dataLabels?: XlsxChartDataLabels;
  series: XlsxChartSeries[];
  seriesNames: string[];
}

export interface XlsxMedia {
  relationshipId: string;
  drawingUri: string;
  targetUri: string;
  type: 'image' | 'embeddedObject';
  name?: string;
  progId?: string;
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
  charts: XlsxChart[];
  media: XlsxMedia[];
  tables: XlsxTable[];
  comments: XlsxComment[];
  threadedComments: XlsxThreadedComment[];
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
  threadedCommentPersons: XlsxThreadedCommentPerson[];
}
