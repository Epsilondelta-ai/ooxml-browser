import { getParsedXmlPart, relationshipById, relationshipsFor, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { WorkbookSheet, WorksheetCell, XlsxCellFormat, XlsxNumberFormat, XlsxStyleTable, XlsxWorkbook } from './model';

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

    return [parseSheet(graph, relationship.resolvedTarget, xmlAttr(sheet, 'name') ?? 'Sheet')];
  });

  return {
    kind: 'xlsx',
    packageGraph: graph,
    sheets,
    sharedStrings,
    styles: parseStyles(graph, workbookUri)
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

function parseSheet(graph: PackageGraph, uri: string, name: string): WorkbookSheet {
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

  return {
    name,
    uri,
    rows
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
