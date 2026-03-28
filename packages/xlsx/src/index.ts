import {
  getParsedXmlPart,
  relationshipById,
  relationshipsFor,
  type PackageGraph
} from '@ooxml/core';
import { xmlAttr, xmlChild, xmlChildren, xmlText } from '@ooxml/core';

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

export function parseXlsx(graph: PackageGraph): XlsxWorkbook {
  const workbookUri = graph.rootDocumentUri ?? '/xl/workbook.xml';
  const workbookXml = getParsedXmlPart(graph, workbookUri);
  if (!workbookXml) {
    throw new Error('Workbook part is missing.');
  }

  const sharedStrings = parseSharedStrings(graph, workbookUri);
  const workbook = workbookXml.document['workbook'];
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
    sharedStrings
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

  const root = xml.document['sst'];
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

  const worksheet = xml.document['worksheet'];
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
