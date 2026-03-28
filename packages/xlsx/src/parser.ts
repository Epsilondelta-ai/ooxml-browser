import { findElementsByLocalName, getParsedXmlPart, relationshipById, relationshipsFor, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { WorkbookSheet, WorksheetCell, XlsxCellFormat, XlsxChart, XlsxComment, XlsxDefinedName, XlsxFrozenPane, XlsxMedia, XlsxNumberFormat, XlsxPageMargins, XlsxPageSetup, XlsxStyleTable, XlsxTable, XlsxWorkbook } from './model';

export function parseXlsx(graph: PackageGraph): XlsxWorkbook {
  const workbookUri = graph.rootDocumentUri ?? '/xl/workbook.xml';
  const workbookXml = getParsedXmlPart(graph, workbookUri);
  if (!workbookXml) {
    throw new Error('Workbook part is missing.');
  }

  const sharedStrings = parseSharedStrings(graph, workbookUri);
  const workbook = workbookXml.document.workbook;
  const sheetsRoot = xmlChild<Record<string, unknown>>(workbook, 'sheets');

  const sheets = xmlChildren<Record<string, unknown>>(sheetsRoot, 'sheet').flatMap((sheet) => {
    const relationshipId = xmlAttr(sheet, 'r:id');
    const relationship = relationshipId ? relationshipById(graph, workbookUri, relationshipId) : undefined;
    if (!relationship?.resolvedTarget) {
      return [];
    }

    return [parseSheet(
      graph,
      relationship.resolvedTarget,
      xmlAttr(sheet, 'name') ?? 'Sheet',
      Number(xmlAttr(sheet, 'sheetId') ?? '0'),
      relationshipId ?? ''
    )];
  });

  return {
    kind: 'xlsx',
    packageGraph: graph,
    sheets,
    sharedStrings,
    styles: parseStyles(graph, workbookUri),
    definedNames: parseDefinedNames(workbook as Record<string, unknown>)
  };
}

function parseSharedStrings(graph: PackageGraph, workbookUri: string): string[] {
  const sharedStringsRelationship = relationshipsFor(graph, workbookUri).find((relationship) => relationship.type.includes('/sharedStrings'));
  if (!sharedStringsRelationship?.resolvedTarget) {
    return [];
  }

  const xml = getParsedXmlPart(graph, sharedStringsRelationship.resolvedTarget);
  if (!xml) {
    return [];
  }

  const root = xml.document.sst;
  return xmlChildren<Record<string, unknown>>(root, 'si').map((item) => {
    const directText = xmlChildren<Record<string, unknown>>(item, 't').map((textNode) => xmlText(textNode)).join('');
    if (directText) {
      return directText;
    }

    return xmlChildren<Record<string, unknown>>(item, 'r').map((run) => xmlText(xmlChild(run, 't'))).join('');
  });
}

function parseSheet(graph: PackageGraph, uri: string, name: string, sheetId: number, relationshipId: string): WorkbookSheet {
  const xml = getParsedXmlPart(graph, uri);
  if (!xml) {
    throw new Error(`Worksheet part ${uri} is missing.`);
  }

  const worksheet = xml.document.worksheet;
  const sheetData = xmlChild<Record<string, unknown>>(worksheet, 'sheetData');
  const rows = xmlChildren<Record<string, unknown>>(sheetData, 'row').map((row) => ({
    index: Number(xmlAttr(row, 'r') ?? '0'),
    cells: xmlChildren<Record<string, unknown>>(row, 'c').map((cell) => parseCell(cell, graph))
  }));

  const mergeCells = xmlChild<Record<string, unknown>>(worksheet, 'mergeCells');
  const mergedRanges = xmlChildren<Record<string, unknown>>(mergeCells, 'mergeCell').map((mergeCell) => xmlAttr(mergeCell, 'ref') ?? '').filter(Boolean);

  const frozenPane = parseFrozenPane(xmlChild<Record<string, unknown>>(worksheet, 'sheetViews'));
  const selection = parseSelection(xmlChild<Record<string, unknown>>(worksheet, 'sheetViews'));
  const pageMargins = parsePageMargins(xmlChild<Record<string, unknown>>(worksheet, 'pageMargins'));
  const pageSetup = parsePageSetup(xmlChild<Record<string, unknown>>(worksheet, 'pageSetup'));

  return {
    name,
    uri,
    sheetId,
    relationshipId,
    rows,
    mergedRanges,
    frozenPane,
    selection,
    pageMargins,
    pageSetup,
    charts: parseSheetCharts(graph, uri),
    media: parseSheetMedia(graph, uri),
    tables: parseSheetTables(graph, uri),
    comments: parseSheetComments(graph, uri)
  };
}

function parseCell(cell: Record<string, unknown>, graph: PackageGraph): WorksheetCell {
  const type = xmlAttr(cell, 't') ?? 'n';
  const rawValue = xmlText(xmlChild(cell, 'v'));
  const formula = xmlText(xmlChild(cell, 'f')) || undefined;
  const styleIndexValue = xmlAttr(cell, 's');
  const styleIndex = styleIndexValue ? Number(styleIndexValue) : undefined;
  let value = rawValue;

  if (type === 'inlineStr') {
    value = xmlText(xmlChild(cell, 'is'));
  }

  if (type === 's') {
    const workbook = graph.rootDocumentUri ?? '/xl/workbook.xml';
    const sharedStrings = parseSharedStrings(graph, workbook);
    value = sharedStrings[Number(rawValue)] ?? rawValue;
  }

  return {
    reference: xmlAttr(cell, 'r') ?? '',
    type,
    value,
    formula,
    styleIndex
  };
}

function parseStyles(graph: PackageGraph, workbookUri: string): XlsxStyleTable {
  const stylesRelationship = relationshipsFor(graph, workbookUri).find((relationship) => relationship.type.includes('/styles'));
  const partUri = stylesRelationship?.resolvedTarget ?? (graph.parts['/xl/styles.xml'] ? '/xl/styles.xml' : undefined);
  if (!partUri) {
    return { numberFormats: {}, cellFormats: {} };
  }

  const xml = getParsedXmlPart(graph, partUri);
  if (!xml) {
    return { partUri, numberFormats: {}, cellFormats: {} };
  }

  const root = xml.document.styleSheet;
  const numberFormats = Object.fromEntries(
    xmlChildren<Record<string, unknown>>(xmlChild<Record<string, unknown>>(root, 'numFmts'), 'numFmt').map((numFmtNode) => {
      const id = Number(xmlAttr(numFmtNode, 'numFmtId') ?? '0');
      const item: XlsxNumberFormat = {
        id,
        code: xmlAttr(numFmtNode, 'formatCode') ?? ''
      };
      return [id, item] satisfies [number, XlsxNumberFormat];
    })
  );

  const cellFormats = Object.fromEntries(
    xmlChildren<Record<string, unknown>>(xmlChild<Record<string, unknown>>(root, 'cellXfs'), 'xf').map((xfNode, index) => {
      const item: XlsxCellFormat = {
        id: index,
        numFmtId: Number(xmlAttr(xfNode, 'numFmtId') ?? '0')
      };
      return [index, item] satisfies [number, XlsxCellFormat];
    })
  );

  return { partUri, numberFormats, cellFormats };
}

export function resolveXlsxCellFormat(workbook: XlsxWorkbook, cell: WorksheetCell): XlsxCellFormat | undefined {
  if (cell.styleIndex === undefined) {
    return undefined;
  }

  return workbook.styles.cellFormats[cell.styleIndex];
}

export function formatXlsxCellValue(workbook: XlsxWorkbook, cell: WorksheetCell): string {
  const style = resolveXlsxCellFormat(workbook, cell);
  const numberFormat = style ? workbook.styles.numberFormats[style.numFmtId] : undefined;

  if (!numberFormat) {
    return cell.value;
  }

  const numericValue = Number(cell.value);
  if (Number.isNaN(numericValue)) {
    return cell.value;
  }

  if (numberFormat.code === '0.00%') {
    return `${(numericValue * 100).toFixed(2)}%`;
  }

  if (numberFormat.code === '0%') {
    return `${Math.round(numericValue * 100)}%`;
  }

  return cell.value;
}

function parseDefinedNames(workbook: Record<string, unknown>): XlsxDefinedName[] {
  const definedNamesRoot = xmlChild<Record<string, unknown>>(workbook, 'definedNames');
  return xmlChildren<Record<string, unknown>>(definedNamesRoot, 'definedName').map((definedName) => ({
    name: xmlAttr(definedName, 'name') ?? '',
    reference: xmlText(definedName),
    scopeSheetId: (() => {
      const localSheetId = xmlAttr(definedName, 'localSheetId');
      return localSheetId ? Number(localSheetId) : undefined;
    })()
  }));
}

function parseFrozenPane(sheetViews: Record<string, unknown> | undefined): XlsxFrozenPane | undefined {
  const sheetView = xmlChild<Record<string, unknown>>(sheetViews, 'sheetView');
  const pane = xmlChild<Record<string, unknown>>(sheetView, 'pane');
  if (!pane) {
    return undefined;
  }

  return {
    xSplit: (() => { const value = xmlAttr(pane, 'xSplit'); return value ? Number(value) : undefined; })(),
    ySplit: (() => { const value = xmlAttr(pane, 'ySplit'); return value ? Number(value) : undefined; })(),
    topLeftCell: xmlAttr(pane, 'topLeftCell'),
    state: xmlAttr(pane, 'state')
  };
}

function parseSelection(sheetViews: Record<string, unknown> | undefined): { activeCell?: string; sqref?: string } | undefined {
  const sheetView = xmlChild<Record<string, unknown>>(sheetViews, 'sheetView');
  const selection = xmlChild<Record<string, unknown>>(sheetView, 'selection');
  if (!selection) {
    return undefined;
  }

  return {
    activeCell: xmlAttr(selection, 'activeCell') ?? undefined,
    sqref: xmlAttr(selection, 'sqref') ?? undefined
  };
}

function parsePageMargins(pageMargins: Record<string, unknown> | undefined): XlsxPageMargins | undefined {
  if (!pageMargins) {
    return undefined;
  }

  return {
    left: parseNumericAttr(pageMargins, 'left'),
    right: parseNumericAttr(pageMargins, 'right'),
    top: parseNumericAttr(pageMargins, 'top'),
    bottom: parseNumericAttr(pageMargins, 'bottom'),
    header: parseNumericAttr(pageMargins, 'header'),
    footer: parseNumericAttr(pageMargins, 'footer')
  };
}

function parsePageSetup(pageSetup: Record<string, unknown> | undefined): XlsxPageSetup | undefined {
  if (!pageSetup) {
    return undefined;
  }

  return {
    orientation: xmlAttr(pageSetup, 'orientation') ?? undefined,
    paperSize: parseNumericAttr(pageSetup, 'paperSize'),
    scale: parseNumericAttr(pageSetup, 'scale'),
    fitToWidth: parseNumericAttr(pageSetup, 'fitToWidth'),
    fitToHeight: parseNumericAttr(pageSetup, 'fitToHeight')
  };
}

function parseNumericAttr(node: Record<string, unknown>, attribute: string): number | undefined {
  const value = xmlAttr(node, attribute);
  return value ? Number(value) : undefined;
}

export function extractFormulaReferences(formula: string): string[] {
  const matches = formula.match(/\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g) ?? [];
  return Array.from(new Set(matches));
}

export function resolveDefinedName(workbook: XlsxWorkbook, name: string): XlsxDefinedName | undefined {
  return workbook.definedNames.find((definedName) => definedName.name === name);
}

function parseSheetTables(graph: PackageGraph, sheetUri: string): XlsxTable[] {
  return relationshipsFor(graph, sheetUri)
    .filter((relationship) => relationship.type.includes('/table') && relationship.resolvedTarget)
    .map((relationship) => {
      const xml = getParsedXmlPart(graph, relationship.resolvedTarget!);
      const root = xml?.document.table as Record<string, unknown> | undefined;
      return {
        name: (root ? xmlAttr(root, 'displayName') ?? xmlAttr(root, 'name') : undefined) ?? relationship.id,
        range: (root ? xmlAttr(root, 'ref') : undefined) ?? '',
        partUri: relationship.resolvedTarget!
      };
    });
}

function parseSheetCharts(graph: PackageGraph, sheetUri: string): XlsxChart[] {
  return relationshipsFor(graph, sheetUri)
    .filter((relationship) => relationship.type.includes('/drawing') && relationship.resolvedTarget)
    .flatMap((relationship) => parseDrawingCharts(graph, relationship.resolvedTarget!));
}

function parseDrawingCharts(graph: PackageGraph, drawingUri: string): XlsxChart[] {
  const xml = getParsedXmlPart(graph, drawingUri);
  if (!xml) {
    return [];
  }

  const drawingRelationships = relationshipsFor(graph, drawingUri);
  const frames = findElementsByLocalName(xml.document, 'graphicFrame');
  return frames.flatMap((frame, drawingNameOccurrence) => {
    const graphicData = findElementsByLocalName(frame, 'graphicData')[0];
    const chartNode = graphicData ? findElementsByLocalName(graphicData, 'chart')[0] : undefined;
    const relationshipId = xmlAttr(chartNode, 'r:id');
    const target = relationshipId ? drawingRelationships.find((entry) => entry.id === relationshipId)?.resolvedTarget : undefined;
    if (!relationshipId || !target) {
      return [];
    }

    const nonVisual = findElementsByLocalName(frame, 'cNvPr')[0];
    const chartXml = getParsedXmlPart(graph, target);
    const chartTitle = chartXml ? findElementsByLocalName(chartXml.document, 't').map((node) => xmlText(node)).join('') : undefined;
    return [{
      relationshipId,
      drawingUri,
      drawingNameOccurrence,
      targetUri: target,
      name: xmlAttr(nonVisual, 'name'),
      title: chartTitle || undefined
    }];
  });
}

function parseSheetMedia(graph: PackageGraph, sheetUri: string): XlsxMedia[] {
  return relationshipsFor(graph, sheetUri)
    .filter((relationship) => relationship.type.includes('/drawing') && relationship.resolvedTarget)
    .flatMap((relationship) => parseDrawingMedia(graph, relationship.resolvedTarget!));
}

function parseDrawingMedia(graph: PackageGraph, drawingUri: string): XlsxMedia[] {
  const xml = getParsedXmlPart(graph, drawingUri);
  if (!xml) {
    return [];
  }

  const drawingRelationships = relationshipsFor(graph, drawingUri);
  const pictures = findElementsByLocalName(xml.document, 'pic');
  return pictures.flatMap((picture) => {
    const blip = findElementsByLocalName(picture, 'blip')[0];
    const relationshipId = xmlAttr(blip, 'r:embed');
    const target = relationshipId ? drawingRelationships.find((entry) => entry.id === relationshipId)?.resolvedTarget : undefined;
    if (!relationshipId || !target) {
      return [];
    }

    const nonVisual = findElementsByLocalName(picture, 'cNvPr')[0];
    return [{
      relationshipId,
      drawingUri,
      targetUri: target,
      type: 'image' as const,
      name: xmlAttr(nonVisual, 'name')
    }];
  });
}

function parseSheetComments(graph: PackageGraph, sheetUri: string): XlsxComment[] {
  const commentsRelationship = relationshipsFor(graph, sheetUri).find((relationship) => relationship.type.includes('/comments'));
  if (!commentsRelationship?.resolvedTarget) {
    return [];
  }

  const xml = getParsedXmlPart(graph, commentsRelationship.resolvedTarget);
  if (!xml) {
    return [];
  }

  const commentsRoot = xml.document.comments;
  const authors = xmlChildren<Record<string, unknown>>(xmlChild<Record<string, unknown>>(commentsRoot, 'authors'), 'author').map((author) => xmlText(author));

  return xmlChildren<Record<string, unknown>>(xmlChild<Record<string, unknown>>(commentsRoot, 'commentList'), 'comment').map((comment) => ({
    reference: xmlAttr(comment, 'ref') ?? '',
    author: (() => {
      const authorId = xmlAttr(comment, 'authorId');
      return authorId ? authors[Number(authorId)] : undefined;
    })(),
    text: xmlChildren<Record<string, unknown>>(comment, 'text').map((textNode) => xmlText(textNode)).join('')
  }));
}
