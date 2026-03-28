import { clonePackageGraph, relationshipsFor, serializePackageGraph, updatePackagePartText, xmlAttr, getParsedXmlPart } from '@ooxml/core';
import type { WorkbookSheet, XlsxComment, XlsxDefinedName, XlsxTable, XlsxWorkbook, WorksheetCell } from '@ooxml/xlsx';

export function serializeXlsx(workbook: XlsxWorkbook): Uint8Array {
  const graph = clonePackageGraph(workbook.packageGraph);
  const sharedStringPool = createSharedStringPool(workbook);
  const sharedStringsUri = '/xl/sharedStrings.xml';
  const hasSharedStringsPart = Boolean(graph.parts[sharedStringsUri]);

  updatePackagePartText(
    graph,
    '/xl/workbook.xml',
    buildWorkbookXml(workbook),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
  );

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

    for (const table of sheet.tables) {
      updatePackagePartText(
        graph,
        table.partUri,
        buildTableXml(table, graph),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
      );
    }

    const commentsRelationship = relationshipsFor(graph, sheet.uri).find((relationship) => relationship.type.includes('/comments'));
    if (commentsRelationship?.resolvedTarget) {
      updatePackagePartText(
        graph,
        commentsRelationship.resolvedTarget,
        buildCommentsXml(sheet.comments),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
      );
    }
  }

  return serializePackageGraph(graph);
}

function buildWorkbookXml(workbook: XlsxWorkbook): string {
  const definedNamesXml = workbook.definedNames.length
    ? `<definedNames>${workbook.definedNames.map(buildDefinedNameXml).join('')}</definedNames>`
    : '';
  const sheetsXml = workbook.sheets.map((sheet) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${sheet.sheetId}" r:id="${escapeXml(sheet.relationshipId)}"/>`).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${definedNamesXml}<sheets>${sheetsXml}</sheets></workbook>`;
}

function buildDefinedNameXml(definedName: XlsxDefinedName): string {
  return `<definedName name="${escapeXml(definedName.name)}"${definedName.scopeSheetId !== undefined ? ` localSheetId="${definedName.scopeSheetId}"` : ''}>${escapeXml(definedName.reference)}</definedName>`;
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
  return `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${values.length}" uniqueCount="${values.length}">${items}</sst>`;
}

function buildWorksheetXml(sheet: WorkbookSheet, sharedStringIndices: Map<string, number>, useSharedStrings: boolean): string {
  const paneXml = sheet.frozenPane
    ? `<sheetViews><sheetView workbookViewId="0"><pane${sheet.frozenPane.xSplit !== undefined ? ` xSplit="${sheet.frozenPane.xSplit}"` : ''}${sheet.frozenPane.ySplit !== undefined ? ` ySplit="${sheet.frozenPane.ySplit}"` : ''}${sheet.frozenPane.topLeftCell ? ` topLeftCell="${escapeXml(sheet.frozenPane.topLeftCell)}"` : ''}${sheet.frozenPane.state ? ` state="${escapeXml(sheet.frozenPane.state)}"` : ''}/></sheetView></sheetViews>`
    : '';
  const rows = sheet.rows.map((row) => `<row r="${row.index}">${row.cells.map((cell) => buildCellXml(cell, sharedStringIndices, useSharedStrings)).join('')}</row>`).join('');
  const mergeCellsXml = sheet.mergedRanges.length
    ? `<mergeCells count="${sheet.mergedRanges.length}">${sheet.mergedRanges.map((range) => `<mergeCell ref="${escapeXml(range)}"/>`).join('')}</mergeCells>`
    : '';

  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${paneXml}<sheetData>${rows}</sheetData>${mergeCellsXml}</worksheet>`;
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

function buildTableXml(table: XlsxTable, graph: ReturnType<typeof clonePackageGraph>): string {
  const xml = getParsedXmlPart(graph, table.partUri);
  const root = xml?.document.table as Record<string, unknown> | undefined;
  const tableId = root ? xmlAttr(root, 'id') ?? '1' : '1';

  return `<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${escapeXml(tableId)}" name="${escapeXml(table.name)}" displayName="${escapeXml(table.name)}" ref="${escapeXml(table.range)}"/>`;
}

function buildCommentsXml(comments: XlsxComment[]): string {
  const authors = Array.from(new Set(comments.map((comment) => comment.author).filter((author): author is string => Boolean(author))));
  const authorIndex = new Map(authors.map((author, index) => [author, index]));
  const commentsXml = comments.map((comment) => `<comment ref="${escapeXml(comment.reference)}"${comment.author ? ` authorId="${authorIndex.get(comment.author) ?? 0}"` : ''}><text><r><t>${escapeXml(comment.text)}</t></r></text></comment>`).join('');
  const authorsXml = authors.length ? `<authors>${authors.map((author) => `<author>${escapeXml(author)}</author>`).join('')}</authors>` : '';
  return `<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${authorsXml}<commentList>${commentsXml}</commentList></comments>`;
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
