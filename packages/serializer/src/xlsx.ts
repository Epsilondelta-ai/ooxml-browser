import { applyXmlPatchPlan, clonePackageGraph, relationshipsFor, serializePackageGraph, updatePackagePartText, xmlAttr, getParsedXmlPart } from '@ooxml/core';
import { parseXlsx, type WorkbookSheet, type WorksheetCell, type XlsxComment, type XlsxDefinedName, type XlsxTable, type XlsxWorkbook } from '@ooxml/xlsx';

export function serializeXlsx(workbook: XlsxWorkbook): Uint8Array {
  const graph = clonePackageGraph(workbook.packageGraph);
  const originalWorkbook = parseXlsx(workbook.packageGraph);
  const sharedStringPool = createSharedStringPool(workbook);
  const sharedStringsUri = '/xl/sharedStrings.xml';
  const hasSharedStringsPart = Boolean(graph.parts[sharedStringsUri]);

  updatePackagePartText(
    graph,
    '/xl/workbook.xml',
    patchWorkbookXml(graph.parts['/xl/workbook.xml']?.text, originalWorkbook, workbook) ?? buildWorkbookXml(workbook),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
  );

  if (hasSharedStringsPart && shouldRewriteSharedStrings(originalWorkbook, workbook)) {
    updatePackagePartText(
      graph,
      sharedStringsUri,
      buildSharedStringsXml(sharedStringPool.values),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
    );
  }

  for (const sheet of workbook.sheets) {
    const originalSheet = originalWorkbook.sheets.find((entry) => entry.uri === sheet.uri);
    updatePackagePartText(
      graph,
      sheet.uri,
      buildWorksheetXml(sheet, sharedStringPool.indexByValue, hasSharedStringsPart, graph.parts[sheet.uri]?.text, originalSheet),
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
        buildCommentsXml(sheet.comments, graph.parts[commentsRelationship.resolvedTarget]?.text),
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

  return `<?xml version="1.0" encoding="UTF-8"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${definedNamesXml}<sheets>${sheetsXml}</sheets></workbook>`;
}

function patchWorkbookXml(existingSource: string | undefined, originalWorkbook: XlsxWorkbook, workbook: XlsxWorkbook): string | undefined {
  if (!existingSource || !canPatchWorkbookDefinedNamesOnly(originalWorkbook, workbook)) {
    return undefined;
  }

  const operations = workbook.definedNames.flatMap((definedName, index) => {
    const originalDefinedName = originalWorkbook.definedNames[index];
    if (!originalDefinedName || definedName.reference === originalDefinedName.reference) {
      return [];
    }

    return [{
      op: 'replaceContainerText' as const,
      tagName: 'definedName',
      occurrence: index,
      newText: definedName.reference
    }];
  });

  return operations.length > 0 ? applyXmlPatchPlan(existingSource, operations) : existingSource;
}

function canPatchWorkbookDefinedNamesOnly(originalWorkbook: XlsxWorkbook, workbook: XlsxWorkbook): boolean {
  if (workbook.sheets.length !== originalWorkbook.sheets.length || workbook.definedNames.length !== originalWorkbook.definedNames.length) {
    return false;
  }

  const sheetsStable = workbook.sheets.every((sheet, index) => {
    const originalSheet = originalWorkbook.sheets[index];
    return Boolean(originalSheet)
      && sheet.name === originalSheet.name
      && sheet.sheetId === originalSheet.sheetId
      && sheet.relationshipId === originalSheet.relationshipId;
  });

  if (!sheetsStable) {
    return false;
  }

  return workbook.definedNames.every((definedName, index) => {
    const originalDefinedName = originalWorkbook.definedNames[index];
    return Boolean(originalDefinedName)
      && definedName.name === originalDefinedName.name
      && definedName.scopeSheetId === originalDefinedName.scopeSheetId;
  });
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
  return `<?xml version="1.0" encoding="UTF-8"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${values.length}" uniqueCount="${values.length}">${items}</sst>`;
}
function shouldRewriteSharedStrings(originalWorkbook: XlsxWorkbook, workbook: XlsxWorkbook): boolean {
  return JSON.stringify(sharedStringSignature(originalWorkbook)) !== JSON.stringify(sharedStringSignature(workbook));
}

function sharedStringSignature(workbook: XlsxWorkbook): Array<{ reference: string; value: string }> {
  return workbook.sheets.flatMap((sheet) =>
    sheet.rows.flatMap((row) =>
      row.cells
        .filter((cell) => shouldUseSharedString(cell))
        .map((cell) => ({ reference: cell.reference, value: cell.value }))
    )
  );
}


function buildWorksheetXml(sheet: WorkbookSheet, sharedStringIndices: Map<string, number>, useSharedStrings: boolean, existingSource?: string, originalSheet?: WorkbookSheet): string {
  if (existingSource && originalSheet && canPatchWorksheet(sheet, originalSheet)) {
    let next = existingSource;
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const operations = [] as Array<Parameters<typeof applyXmlPatchPlan>[1][number]>;
        if (cell.formula) {
          operations.push({ op: 'replaceText', containerTag: 'c', keyAttr: 'r', keyValue: cell.reference, textTag: 'f', newText: cell.formula });
        }
        operations.push({
          op: 'replaceText',
          containerTag: 'c',
          keyAttr: 'r',
          keyValue: cell.reference,
          textTag: shouldUseSharedString(cell) && !useSharedStrings ? 't' : 'v',
          newText: cell.value
        });
        next = applyXmlPatchPlan(next, operations);
      }
    }

    if (sheet.frozenPane) {
      const operations = [] as Array<Parameters<typeof applyXmlPatchPlan>[1][number]>;
      if (sheet.frozenPane.topLeftCell) {
        operations.push({ op: 'replaceAttribute', tagName: 'pane', targetAttr: 'topLeftCell', newValue: sheet.frozenPane.topLeftCell });
      }
      if (sheet.frozenPane.xSplit !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pane', targetAttr: 'xSplit', newValue: String(sheet.frozenPane.xSplit) });
      }
      if (sheet.frozenPane.ySplit !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pane', targetAttr: 'ySplit', newValue: String(sheet.frozenPane.ySplit) });
      }
      if (sheet.frozenPane.state) {
        operations.push({ op: 'replaceAttribute', tagName: 'pane', targetAttr: 'state', newValue: sheet.frozenPane.state });
      }
      if (operations.length > 0) {
        next = applyXmlPatchPlan(next, operations);
      }
    }

    return next;
  }

  const paneXml = sheet.frozenPane
    ? `<sheetViews><sheetView workbookViewId="0"><pane${sheet.frozenPane.xSplit !== undefined ? ` xSplit="${sheet.frozenPane.xSplit}"` : ''}${sheet.frozenPane.ySplit !== undefined ? ` ySplit="${sheet.frozenPane.ySplit}"` : ''}${sheet.frozenPane.topLeftCell ? ` topLeftCell="${escapeXml(sheet.frozenPane.topLeftCell)}"` : ''}${sheet.frozenPane.state ? ` state="${escapeXml(sheet.frozenPane.state)}"` : ''}/></sheetView></sheetViews>`
    : '';
  const rows = sheet.rows.map((row) => `<row r="${row.index}">${row.cells.map((cell) => buildCellXml(cell, sharedStringIndices, useSharedStrings)).join('')}</row>`).join('');
  const mergeCellsXml = sheet.mergedRanges.length
    ? `<mergeCells count="${sheet.mergedRanges.length}">${sheet.mergedRanges.map((range) => `<mergeCell ref="${escapeXml(range)}"/>`).join('')}</mergeCells>`
    : '';

  const worksheetOpenTag = preserveWorksheetOpenTag(existingSource) ?? '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  return `<?xml version="1.0" encoding="UTF-8"?>\n${worksheetOpenTag}${paneXml}<sheetData>${rows}</sheetData>${mergeCellsXml}</worksheet>`;
}

function preserveWorksheetOpenTag(existingSource?: string): string | undefined {
  return existingSource?.match(/<worksheet\b[^>]*>/)?.[0];
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
  const source = xml?.source;
  if (source) {
    return applyXmlPatchPlan(source, [
      { op: 'replaceAttribute', tagName: 'table', targetAttr: 'name', newValue: table.name },
      { op: 'replaceAttribute', tagName: 'table', targetAttr: 'displayName', newValue: table.name },
      { op: 'replaceAttribute', tagName: 'table', targetAttr: 'ref', newValue: table.range }
    ]);
  }

  const root = xml?.document.table as Record<string, unknown> | undefined;
  const tableId = root ? xmlAttr(root, 'id') ?? '1' : '1';
  return `<?xml version="1.0" encoding="UTF-8"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${escapeXml(tableId)}" name="${escapeXml(table.name)}" displayName="${escapeXml(table.name)}" ref="${escapeXml(table.range)}"/>`;
}

function buildCommentsXml(comments: XlsxComment[], existingSource?: string): string {
  const normalizedAuthors = comments.map((comment) => comment.author ?? '');
  const authorSetChanged = JSON.stringify(normalizedAuthors) !== JSON.stringify(parseCommentAuthors(existingSource));

  if (existingSource && !authorSetChanged) {
    return applyXmlPatchPlan(existingSource, comments.map((comment) => ({
      op: 'replaceText' as const,
      containerTag: 'comment',
      keyAttr: 'ref',
      keyValue: comment.reference,
      textTag: 't',
      newText: comment.text
    })));
  }

  const authors = Array.from(new Set(comments.map((comment) => comment.author).filter((author): author is string => Boolean(author))));
  const authorIndex = new Map(authors.map((author, index) => [author, index]));
  const commentsXml = comments.map((comment) => `<comment ref="${escapeXml(comment.reference)}"${comment.author ? ` authorId="${authorIndex.get(comment.author) ?? 0}"` : ''}><text><r><t>${escapeXml(comment.text)}</t></r></text></comment>`).join('');
  const authorsXml = authors.length ? `<authors>${authors.map((author) => `<author>${escapeXml(author)}</author>`).join('')}</authors>` : '';
  return `<?xml version="1.0" encoding="UTF-8"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${authorsXml}<commentList>${commentsXml}</commentList></comments>`;
}

function parseCommentAuthors(existingSource?: string): string[] {
  if (!existingSource) {
    return [];
  }

  const authorPattern = /<comment\b[^>]*?(?:authorId="(\d+)")?[^>]*>/g;
  const authorsRoot = existingSource.match(/<authors>([\s\S]*?)<\/authors>/)?.[1] ?? '';
  const authorValues = [...authorsRoot.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((match) => decodeXml(match[1] ?? ''));

  return [...existingSource.matchAll(authorPattern)].map((match) => {
    const index = match[1] !== undefined ? Number(match[1]) : undefined;
    return index !== undefined && authorValues[index] !== undefined ? authorValues[index] : '';
  });
}

function decodeXml(value: string): string {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');
}

function canPatchWorksheet(sheet: WorkbookSheet, originalSheet: WorkbookSheet): boolean {
  return sheet.rows.every((row) => row.cells.length > 0)
    && sheet.rows.length === originalSheet.rows.length
    && JSON.stringify(sheet.mergedRanges) === JSON.stringify(originalSheet.mergedRanges);
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
