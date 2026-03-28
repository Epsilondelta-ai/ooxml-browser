import { clonePackageGraph, serializePackageGraph, updatePackagePartText } from '@ooxml/core';
import type { WorkbookSheet, WorksheetCell, XlsxWorkbook } from '@ooxml/xlsx';

export function serializeXlsx(workbook: XlsxWorkbook): Uint8Array {
  const graph = clonePackageGraph(workbook.packageGraph);
  const sharedStringPool = createSharedStringPool(workbook);
  const sharedStringsUri = '/xl/sharedStrings.xml';

  const hasSharedStringsPart = Boolean(graph.parts[sharedStringsUri]);

  if (hasSharedStringsPart) {
    updatePackagePartText(
      graph,
      sharedStringsUri,
      buildSharedStringsXml(sharedStringPool.values),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
    );
  }

  for (const sheet of workbook.sheets) {
    updatePackagePartText(
      graph,
      sheet.uri,
      buildWorksheetXml(sheet, sharedStringPool.indexByValue, hasSharedStringsPart),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
    );
  }

  return serializePackageGraph(graph);
}

function createSharedStringPool(workbook: XlsxWorkbook): { values: string[]; indexByValue: Map<string, number> } {
  const values: string[] = [];
  const indexByValue = new Map<string, number>();

  for (const sheet of workbook.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (shouldUseSharedString(cell)) {
          if (!indexByValue.has(cell.value)) {
            indexByValue.set(cell.value, values.length);
            values.push(cell.value);
          }
        }
      }
    }
  }

  return { values, indexByValue };
}

function buildSharedStringsXml(values: string[]): string {
  const items = values.map((value) => `<si><t>${escapeXml(value)}</t></si>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${values.length}" uniqueCount="${values.length}">${items}</sst>`;
}

function buildWorksheetXml(sheet: WorkbookSheet, sharedStringIndices: Map<string, number>, useSharedStrings: boolean): string {
  const rows = sheet.rows.map((row) => `<row r="${row.index}">${row.cells.map((cell) => buildCellXml(cell, sharedStringIndices, useSharedStrings)).join('')}</row>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`;
}

function buildCellXml(cell: WorksheetCell, sharedStringIndices: Map<string, number>, useSharedStrings: boolean): string {
  const styleAttribute = cell.styleIndex !== undefined ? ` s="${cell.styleIndex}"` : '';

  if (cell.formula) {
    return `<c r="${escapeXml(cell.reference)}"${styleAttribute}><f>${escapeXml(cell.formula)}</f><v>${escapeXml(cell.value)}</v></c>`;
  }

  if (shouldUseSharedString(cell)) {
    if (useSharedStrings) {
      const sharedIndex = sharedStringIndices.get(cell.value) ?? 0;
      return `<c r="${escapeXml(cell.reference)}" t="s"${styleAttribute}><v>${sharedIndex}</v></c>`;
    }

    return `<c r="${escapeXml(cell.reference)}" t="inlineStr"${styleAttribute}><is><t>${escapeXml(cell.value)}</t></is></c>`;
  }

  return `<c r="${escapeXml(cell.reference)}"${cell.type !== 'n' ? ` t="${escapeXml(cell.type)}"` : ''}${styleAttribute}><v>${escapeXml(cell.value)}</v></c>`;
}

function shouldUseSharedString(cell: WorksheetCell): boolean {
  if (cell.type === 's') {
    return true;
  }

  return cell.type !== 'n' && !cell.formula && Number.isNaN(Number(cell.value));
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
