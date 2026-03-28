import { applyXmlPatchPlan, clonePackageGraph, relationshipsFor, serializePackageGraph, setRelationshipsForSource, updatePackagePartText, xmlAttr, getParsedXmlPart, upsertRelationship } from '@ooxml/core';
import { parseXlsx, type WorkbookSheet, type WorksheetCell, type XlsxComment, type XlsxDefinedName, type XlsxTable, type XlsxThreadedComment, type XlsxThreadedCommentPerson, type XlsxWorkbook } from '@ooxml/xlsx';

export function serializeXlsx(workbook: XlsxWorkbook): Uint8Array {
  const graph = clonePackageGraph(workbook.packageGraph);
  const originalWorkbook = parseXlsx(workbook.packageGraph);
  const sharedStringPool = createSharedStringPool(workbook);
  const sharedStringsUri = '/xl/sharedStrings.xml';
  const hasSharedStringsPart = Boolean(graph.parts[sharedStringsUri]);

  updatePackagePartText(
    graph,
    '/xl/workbook.xml',
    patchWorkbookXml(graph.parts['/xl/workbook.xml']?.text, originalWorkbook, workbook) ?? buildWorkbookXml(workbook, graph.parts['/xl/workbook.xml']?.text),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
  );

  syncWorkbookThreadedCommentPersons(graph, workbook);

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
    syncWorksheetTableRelationships(graph, originalSheet, sheet);
    syncWorksheetChartRelationships(graph, sheet);
    syncWorksheetChartParts(graph, originalSheet, sheet);
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
    if (sheet.comments.length > 0) {
      const commentsUri = commentsRelationship?.resolvedTarget ?? ensureWorksheetCommentsPart(graph, sheet.uri);
      if (!commentsUri) {
        continue;
      }
      updatePackagePartText(
        graph,
        commentsUri,
        buildCommentsXml(sheet.comments, graph.parts[commentsUri]?.text),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
      );
    }
    if (sheet.comments.length === 0 && commentsRelationship?.resolvedTarget) {
      updatePackagePartText(
        graph,
        commentsRelationship.resolvedTarget,
        buildCommentsXml([], graph.parts[commentsRelationship.resolvedTarget]?.text),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
      );
    }

    const threadedCommentsRelationship = relationshipsFor(graph, sheet.uri).find((relationship) => relationship.type.includes('/threadedComment'));
    if (sheet.threadedComments.length > 0) {
      const threadedCommentsUri = threadedCommentsRelationship?.resolvedTarget ?? ensureWorksheetThreadedCommentsPart(graph, sheet.uri);
      if (!threadedCommentsUri) {
        continue;
      }
      updatePackagePartText(
        graph,
        threadedCommentsUri,
        buildThreadedCommentsXml(sheet.threadedComments),
        'application/vnd.ms-excel.threadedcomments+xml'
      );
    }
    if (sheet.threadedComments.length === 0 && threadedCommentsRelationship?.resolvedTarget) {
      updatePackagePartText(
        graph,
        threadedCommentsRelationship.resolvedTarget,
        buildThreadedCommentsXml([]),
        'application/vnd.ms-excel.threadedcomments+xml'
      );
    }
  }

  return serializePackageGraph(graph);
}

function syncWorkbookThreadedCommentPersons(graph: XlsxWorkbook['packageGraph'], workbook: XlsxWorkbook): void {
  const workbookUri = workbook.packageGraph.rootDocumentUri ?? '/xl/workbook.xml';
  const workbookRelationships = relationshipsFor(graph, workbookUri);
  const peopleRelationship = workbookRelationships.find((relationship) => relationship.type.includes('/person'));
  if (workbook.threadedCommentPersons.length === 0 && !peopleRelationship) {
    return;
  }

  const peopleUri = peopleRelationship?.resolvedTarget ?? ensureWorkbookThreadedCommentPersonsPart(graph, workbookUri);
  if (!peopleUri) {
    return;
  }

  updatePackagePartText(
    graph,
    peopleUri,
    buildThreadedCommentPersonsXml(workbook.threadedCommentPersons),
    'application/vnd.ms-excel.person+xml'
  );
}

function syncWorksheetChartRelationships(graph: XlsxWorkbook['packageGraph'], sheet: WorkbookSheet): void {
  for (const chart of sheet.charts) {
    upsertRelationship(graph, chart.drawingUri, {
      id: chart.relationshipId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
      target: relativeRelationshipTarget(chart.drawingUri, chart.targetUri),
      targetMode: 'Internal'
    });
  }

  for (const media of sheet.media) {
    upsertRelationship(graph, media.drawingUri, {
      id: media.relationshipId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: relativeRelationshipTarget(media.drawingUri, media.targetUri),
      targetMode: 'Internal'
    });
  }
}

function syncWorksheetChartParts(graph: XlsxWorkbook['packageGraph'], originalSheet: WorkbookSheet | undefined, sheet: WorkbookSheet): void {
  for (const chart of sheet.charts) {
    const originalChart = originalSheet?.charts.find((entry) => entry.relationshipId === chart.relationshipId && entry.drawingUri === chart.drawingUri);
    if (
      chart.chartType !== originalChart?.chartType
      || chart.varyColors !== originalChart?.varyColors
      || chart.gapWidth !== originalChart?.gapWidth
      || chart.title !== originalChart?.title
      || chart.legendPosition !== originalChart?.legendPosition
      || chart.categoryAxisTitle !== originalChart?.categoryAxisTitle
      || chart.categoryAxisPosition !== originalChart?.categoryAxisPosition
      || chart.valueAxisTitle !== originalChart?.valueAxisTitle
      || chart.valueAxisPosition !== originalChart?.valueAxisPosition
      || JSON.stringify(chart.dataLabels ?? null) !== JSON.stringify(originalChart?.dataLabels ?? null)
      || JSON.stringify(chart.series) !== JSON.stringify(originalChart?.series ?? [])
    ) {
      const existingSource = chart.chartType === originalChart?.chartType ? graph.parts[chart.targetUri]?.text : undefined;
      updatePackagePartText(
        graph,
        chart.targetUri,
        buildChartXml(chart, existingSource),
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
      );
    }
    if (chart.name !== originalChart?.name) {
      const drawingSource = graph.parts[chart.drawingUri]?.text;
      if (drawingSource) {
        updatePackagePartText(
          graph,
          chart.drawingUri,
          applyXmlPatchPlan(drawingSource, [
            {
              op: 'replaceAttribute',
              tagName: 'xdr:cNvPr',
              targetAttr: 'name',
              newValue: chart.name ?? '',
              occurrence: chart.drawingNameOccurrence
            }
          ]),
          'application/vnd.openxmlformats-officedocument.drawing+xml'
        );
      }
    }
  }
}

function syncWorksheetTableRelationships(graph: XlsxWorkbook['packageGraph'], originalSheet: WorkbookSheet | undefined, sheet: WorkbookSheet): void {
  if (!originalSheet) {
    return;
  }

  const retainedPartUris = new Set(sheet.tables.map((table) => table.partUri));
  const existingRelationships = graph.relationshipsBySource[sheet.uri] ?? [];
  const nextRelationships = existingRelationships.filter((relationship) => {
    if (!relationship.type.includes('/table') || !relationship.resolvedTarget) {
      return true;
    }

    return retainedPartUris.has(relationship.resolvedTarget);
  });

  if (nextRelationships.length !== existingRelationships.length) {
    setRelationshipsForSource(graph, sheet.uri, nextRelationships);
  }
}

function buildWorkbookXml(workbook: XlsxWorkbook, existingSource?: string): string {
  const definedNamesXml = workbook.definedNames.length
    ? `<definedNames>${workbook.definedNames.map(buildDefinedNameXml).join('')}</definedNames>`
    : '';
  const sheetsXml = workbook.sheets.map((sheet) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${sheet.sheetId}" r:id="${escapeXml(sheet.relationshipId)}"/>`).join('');

  const workbookOpenTag = preserveWorkbookOpenTag(existingSource) ?? '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
  return `<?xml version="1.0" encoding="UTF-8"?>\n${workbookOpenTag}${definedNamesXml}<sheets>${sheetsXml}</sheets></workbook>`;
}

function preserveWorkbookOpenTag(existingSource?: string): string | undefined {
  return existingSource?.match(/<workbook\b[^>]*>/)?.[0];
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
        if (cell.styleIndex !== undefined) {
          operations.push({
            op: 'replaceAttribute',
            tagName: 'c',
            keyAttr: 'r',
            keyValue: cell.reference,
            targetAttr: 's',
            newValue: String(cell.styleIndex)
          });
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

    if (sheet.selection) {
      const operations = [] as Array<Parameters<typeof applyXmlPatchPlan>[1][number]>;
      if (sheet.selection.activeCell) {
        operations.push({ op: 'replaceAttribute', tagName: 'selection', targetAttr: 'activeCell', newValue: sheet.selection.activeCell });
      }
      if (sheet.selection.sqref) {
        operations.push({ op: 'replaceAttribute', tagName: 'selection', targetAttr: 'sqref', newValue: sheet.selection.sqref });
      }
      if (operations.length > 0) {
        next = applyXmlPatchPlan(next, operations);
      }
    }

    if (sheet.pageMargins) {
      const operations = [] as Array<Parameters<typeof applyXmlPatchPlan>[1][number]>;
      if (sheet.pageMargins.left !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageMargins', targetAttr: 'left', newValue: String(sheet.pageMargins.left) });
      }
      if (sheet.pageMargins.right !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageMargins', targetAttr: 'right', newValue: String(sheet.pageMargins.right) });
      }
      if (sheet.pageMargins.top !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageMargins', targetAttr: 'top', newValue: String(sheet.pageMargins.top) });
      }
      if (sheet.pageMargins.bottom !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageMargins', targetAttr: 'bottom', newValue: String(sheet.pageMargins.bottom) });
      }
      if (sheet.pageMargins.header !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageMargins', targetAttr: 'header', newValue: String(sheet.pageMargins.header) });
      }
      if (sheet.pageMargins.footer !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageMargins', targetAttr: 'footer', newValue: String(sheet.pageMargins.footer) });
      }
      if (operations.length > 0) {
        next = applyXmlPatchPlan(next, operations);
      }
    }

    if (sheet.pageSetup) {
      const operations = [] as Array<Parameters<typeof applyXmlPatchPlan>[1][number]>;
      if (sheet.pageSetup.orientation) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageSetup', targetAttr: 'orientation', newValue: sheet.pageSetup.orientation });
      }
      if (sheet.pageSetup.paperSize !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageSetup', targetAttr: 'paperSize', newValue: String(sheet.pageSetup.paperSize) });
      }
      if (sheet.pageSetup.scale !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageSetup', targetAttr: 'scale', newValue: String(sheet.pageSetup.scale) });
      }
      if (sheet.pageSetup.fitToWidth !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageSetup', targetAttr: 'fitToWidth', newValue: String(sheet.pageSetup.fitToWidth) });
      }
      if (sheet.pageSetup.fitToHeight !== undefined) {
        operations.push({ op: 'replaceAttribute', tagName: 'pageSetup', targetAttr: 'fitToHeight', newValue: String(sheet.pageSetup.fitToHeight) });
      }
      if (operations.length > 0) {
        next = applyXmlPatchPlan(next, operations);
      }
    }

    return next;
  }

  const paneXml = sheet.frozenPane
    ? `<pane${sheet.frozenPane.xSplit !== undefined ? ` xSplit="${sheet.frozenPane.xSplit}"` : ''}${sheet.frozenPane.ySplit !== undefined ? ` ySplit="${sheet.frozenPane.ySplit}"` : ''}${sheet.frozenPane.topLeftCell ? ` topLeftCell="${escapeXml(sheet.frozenPane.topLeftCell)}"` : ''}${sheet.frozenPane.state ? ` state="${escapeXml(sheet.frozenPane.state)}"` : ''}/>`
    : '';
  const selectionXml = sheet.selection
    ? `<selection${sheet.selection.activeCell ? ` activeCell="${escapeXml(sheet.selection.activeCell)}"` : ''}${sheet.selection.sqref ? ` sqref="${escapeXml(sheet.selection.sqref)}"` : ''}/>`
    : '';
  const sheetViewsXml = paneXml || selectionXml
    ? `<sheetViews><sheetView workbookViewId="0">${paneXml}${selectionXml}</sheetView></sheetViews>`
    : '';
  const rows = sheet.rows.map((row) => `<row r="${row.index}">${row.cells.map((cell) => buildCellXml(cell, sharedStringIndices, useSharedStrings)).join('')}</row>`).join('');
  const mergeCellsXml = sheet.mergedRanges.length
    ? `<mergeCells count="${sheet.mergedRanges.length}">${sheet.mergedRanges.map((range) => `<mergeCell ref="${escapeXml(range)}"/>`).join('')}</mergeCells>`
    : '';
  const pageMarginsXml = sheet.pageMargins
    ? `<pageMargins${sheet.pageMargins.left !== undefined ? ` left="${sheet.pageMargins.left}"` : ''}${sheet.pageMargins.right !== undefined ? ` right="${sheet.pageMargins.right}"` : ''}${sheet.pageMargins.top !== undefined ? ` top="${sheet.pageMargins.top}"` : ''}${sheet.pageMargins.bottom !== undefined ? ` bottom="${sheet.pageMargins.bottom}"` : ''}${sheet.pageMargins.header !== undefined ? ` header="${sheet.pageMargins.header}"` : ''}${sheet.pageMargins.footer !== undefined ? ` footer="${sheet.pageMargins.footer}"` : ''}/>`
    : '';
  const pageSetupXml = sheet.pageSetup
    ? `<pageSetup${sheet.pageSetup.orientation ? ` orientation="${escapeXml(sheet.pageSetup.orientation)}"` : ''}${sheet.pageSetup.paperSize !== undefined ? ` paperSize="${sheet.pageSetup.paperSize}"` : ''}${sheet.pageSetup.scale !== undefined ? ` scale="${sheet.pageSetup.scale}"` : ''}${sheet.pageSetup.fitToWidth !== undefined ? ` fitToWidth="${sheet.pageSetup.fitToWidth}"` : ''}${sheet.pageSetup.fitToHeight !== undefined ? ` fitToHeight="${sheet.pageSetup.fitToHeight}"` : ''}/>`
    : '';

  const worksheetOpenTag = preserveWorksheetOpenTag(existingSource) ?? '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  return `<?xml version="1.0" encoding="UTF-8"?>\n${worksheetOpenTag}${sheetViewsXml}<sheetData>${rows}</sheetData>${mergeCellsXml}${pageMarginsXml}${pageSetupXml}</worksheet>`;
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

function buildChartXml(chart: WorkbookSheet['charts'][number], existingSource?: string): string {
    if (existingSource) {
      const operations: Array<Parameters<typeof applyXmlPatchPlan>[1][number]> = [];
      if (chart.title !== undefined) {
      operations.push({
        op: 'replaceText',
        containerTag: 'c:title',
        occurrence: 0,
        textTag: 'a:t',
          newText: chart.title
        });
      }
      if (chart.varyColors !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:varyColors',
          targetAttr: 'val',
          newValue: chart.varyColors ? '1' : '0'
        });
      }
      if (chart.gapWidth !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:gapWidth',
          targetAttr: 'val',
          newValue: String(chart.gapWidth)
        });
      }
      if (chart.legendPosition !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:legendPos',
          targetAttr: 'val',
          newValue: chart.legendPosition
        });
      }
      if (chart.categoryAxisTitle !== undefined) {
        operations.push({
          op: 'replaceText',
          containerTag: 'c:catAx',
          occurrence: 0,
          textTag: 'a:t',
          newText: chart.categoryAxisTitle
        });
      }
      if (chart.categoryAxisPosition !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:axPos',
          targetAttr: 'val',
          newValue: chart.categoryAxisPosition,
          occurrence: 0
        });
      }
      if (chart.valueAxisTitle !== undefined) {
        operations.push({
          op: 'replaceText',
          containerTag: 'c:valAx',
          occurrence: 0,
          textTag: 'a:t',
          newText: chart.valueAxisTitle
        });
      }
      if (chart.valueAxisPosition !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:axPos',
          targetAttr: 'val',
          newValue: chart.valueAxisPosition,
          occurrence: 1
        });
      }
      if (chart.dataLabels?.position !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:dLblPos',
          targetAttr: 'val',
          newValue: chart.dataLabels.position
        });
      }
      if (chart.dataLabels?.showValue !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:showVal',
          targetAttr: 'val',
          newValue: chart.dataLabels.showValue ? '1' : '0'
        });
      }
      if (chart.dataLabels?.showCategoryName !== undefined) {
        operations.push({
          op: 'replaceAttribute',
          tagName: 'c:showCatName',
          targetAttr: 'val',
          newValue: chart.dataLabels.showCategoryName ? '1' : '0'
        });
      }
      for (const [seriesIndex, series] of chart.series.entries()) {
        operations.push({
          op: 'replaceText',
          containerTag: 'c:ser',
          occurrence: seriesIndex,
          textTag: 'a:t',
          newText: series.name
        });
        if (series.invertIfNegative !== undefined) {
          operations.push({
            op: 'replaceAttribute',
            tagName: 'c:invertIfNegative',
            targetAttr: 'val',
            newValue: series.invertIfNegative ? '1' : '0',
            occurrence: seriesIndex
          });
        }
      }
    if (operations.length > 0) {
      return applyXmlPatchPlan(existingSource, operations);
    }
  }

  const chartType = chart.chartType ?? 'barChart';
  const seriesXml = chart.series.map((seriesEntry, index) => `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:rich><a:t>${escapeXml(seriesEntry.name)}</a:t></c:rich></c:tx>${seriesEntry.invertIfNegative !== undefined ? `<c:invertIfNegative val="${seriesEntry.invertIfNegative ? '1' : '0'}"/>` : ''}</c:ser>`).join('');
  const varyColorsXml = chart.varyColors !== undefined ? `<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>` : '';
  const gapWidthXml = chart.gapWidth !== undefined ? `<c:gapWidth val="${chart.gapWidth}"/>` : '';
  const categoryAxisXml = chart.categoryAxisTitle || chart.categoryAxisPosition ? `<c:catAx>${chart.categoryAxisTitle ? `<c:title><c:tx><c:rich><a:t>${escapeXml(chart.categoryAxisTitle)}</a:t></c:rich></c:tx></c:title>` : ''}${chart.categoryAxisPosition ? `<c:axPos val="${escapeXml(chart.categoryAxisPosition)}"/>` : ''}</c:catAx>` : '';
  const valueAxisXml = chart.valueAxisTitle || chart.valueAxisPosition ? `<c:valAx>${chart.valueAxisTitle ? `<c:title><c:tx><c:rich><a:t>${escapeXml(chart.valueAxisTitle)}</a:t></c:rich></c:tx></c:title>` : ''}${chart.valueAxisPosition ? `<c:axPos val="${escapeXml(chart.valueAxisPosition)}"/>` : ''}</c:valAx>` : '';
  const dataLabelsXml = chart.dataLabels ? `<c:dLbls>${chart.dataLabels.position ? `<c:dLblPos val="${escapeXml(chart.dataLabels.position)}"/>` : ''}${chart.dataLabels.showValue !== undefined ? `<c:showVal val="${chart.dataLabels.showValue ? '1' : '0'}"/>` : ''}${chart.dataLabels.showCategoryName !== undefined ? `<c:showCatName val="${chart.dataLabels.showCategoryName ? '1' : '0'}"/>` : ''}</c:dLbls>` : '';
  const legendXml = chart.legendPosition ? `<c:legend><c:legendPos val="${escapeXml(chart.legendPosition)}"/></c:legend>` : '';
  return `<?xml version="1.0" encoding="UTF-8"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:t>${escapeXml(chart.title ?? '')}</a:t></c:rich></c:tx></c:title><c:plotArea><c:${chartType}>${varyColorsXml}${seriesXml}${dataLabelsXml}${gapWidthXml}</c:${chartType}>${categoryAxisXml}${valueAxisXml}</c:plotArea>${legendXml}</c:chart></c:chartSpace>`;
}

function buildCommentsXml(comments: XlsxComment[], existingSource?: string): string {
  if (comments.length === 0) {
    return `<?xml version="1.0" encoding="UTF-8"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><commentList></commentList></comments>`;
  }

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

function buildThreadedCommentsXml(comments: XlsxThreadedComment[]): string {
  const commentsXml = comments.map((comment) => `<threadedComment ref="${escapeXml(comment.reference)}" personId="${escapeXml(comment.personId)}" id="${escapeXml(comment.id)}"${comment.parentId ? ` parentId="${escapeXml(comment.parentId)}"` : ''}><text>${escapeXml(comment.text)}</text></threadedComment>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">${commentsXml}</ThreadedComments>`;
}

function buildThreadedCommentPersonsXml(persons: XlsxThreadedCommentPerson[]): string {
  const body = persons.map((person) => `<person id="${escapeXml(person.id)}" displayName="${escapeXml(person.displayName)}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/person">${body}</personList>`;
}

function ensureWorksheetCommentsPart(graph: XlsxWorkbook['packageGraph'], sheetUri: string): string | undefined {
  const sheetRelationships = relationshipsFor(graph, sheetUri);
  const existingComments = sheetRelationships.find((relationship) => relationship.type.includes('/comments'));
  if (existingComments?.resolvedTarget) {
    return existingComments.resolvedTarget;
  }

  const commentsUri = nextCommentsUri(graph.parts);
  const relationshipId = nextRelationshipId(sheetRelationships);
  updatePackagePartText(
    graph,
    commentsUri,
    `<?xml version="1.0" encoding="UTF-8"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><commentList></commentList></comments>`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
  );
  upsertRelationship(graph, sheetUri, {
    id: relationshipId,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    target: relativeRelationshipTarget(sheetUri, commentsUri),
    targetMode: 'Internal'
  });
  return commentsUri;
}

function ensureWorksheetThreadedCommentsPart(graph: XlsxWorkbook['packageGraph'], sheetUri: string): string | undefined {
  const sheetRelationships = relationshipsFor(graph, sheetUri);
  const existingThreadedComments = sheetRelationships.find((relationship) => relationship.type.includes('/threadedComment'));
  if (existingThreadedComments?.resolvedTarget) {
    return existingThreadedComments.resolvedTarget;
  }

  const threadedCommentsUri = nextThreadedCommentsUri(graph.parts);
  const relationshipId = nextRelationshipId(sheetRelationships);
  updatePackagePartText(
    graph,
    threadedCommentsUri,
    buildThreadedCommentsXml([]),
    'application/vnd.ms-excel.threadedcomments+xml'
  );
  upsertRelationship(graph, sheetUri, {
    id: relationshipId,
    type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
    target: relativeRelationshipTarget(sheetUri, threadedCommentsUri),
    targetMode: 'Internal'
  });
  return threadedCommentsUri;
}

function ensureWorkbookThreadedCommentPersonsPart(graph: XlsxWorkbook['packageGraph'], workbookUri: string): string | undefined {
  const workbookRelationships = relationshipsFor(graph, workbookUri);
  const existingPeople = workbookRelationships.find((relationship) => relationship.type.includes('/person'));
  if (existingPeople?.resolvedTarget) {
    return existingPeople.resolvedTarget;
  }

  const peopleUri = nextThreadedCommentPersonsUri(graph.parts);
  const relationshipId = nextRelationshipId(workbookRelationships);
  updatePackagePartText(
    graph,
    peopleUri,
    buildThreadedCommentPersonsXml([]),
    'application/vnd.ms-excel.person+xml'
  );
  upsertRelationship(graph, workbookUri, {
    id: relationshipId,
    type: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
    target: relativeRelationshipTarget(workbookUri, peopleUri),
    targetMode: 'Internal'
  });
  return peopleUri;
}

function nextCommentsUri(parts: XlsxWorkbook['packageGraph']['parts']): string {
  let candidateIndex = 1;
  let candidate = `/xl/comments${candidateIndex}.xml`;
  while (parts[candidate]) {
    candidateIndex += 1;
    candidate = `/xl/comments${candidateIndex}.xml`;
  }
  return candidate;
}

function nextThreadedCommentsUri(parts: XlsxWorkbook['packageGraph']['parts']): string {
  let candidateIndex = 1;
  let candidate = `/xl/threadedComments/threadedComment${candidateIndex}.xml`;
  while (parts[candidate]) {
    candidateIndex += 1;
    candidate = `/xl/threadedComments/threadedComment${candidateIndex}.xml`;
  }
  return candidate;
}

function nextThreadedCommentPersonsUri(parts: XlsxWorkbook['packageGraph']['parts']): string {
  let candidateIndex = 1;
  let candidate = `/xl/persons/person${candidateIndex}.xml`;
  while (parts[candidate]) {
    candidateIndex += 1;
    candidate = `/xl/persons/person${candidateIndex}.xml`;
  }
  return candidate;
}

function nextRelationshipId(relationships: ReturnType<typeof relationshipsFor>): string {
  let candidateIndex = relationships.length + 1;
  let candidate = `rId${candidateIndex}`;
  const existingIds = new Set(relationships.map((relationship) => relationship.id));
  while (existingIds.has(candidate)) {
    candidateIndex += 1;
    candidate = `rId${candidateIndex}`;
  }
  return candidate;
}

function relativeRelationshipTarget(sourceUri: string, targetUri: string): string {
  const sourceSegments = sourceUri.replace(/^\//, '').split('/');
  sourceSegments.pop();
  const targetSegments = targetUri.replace(/^\//, '').split('/');

  while (sourceSegments.length > 0 && targetSegments.length > 0 && sourceSegments[0] === targetSegments[0]) {
    sourceSegments.shift();
    targetSegments.shift();
  }

  return `${sourceSegments.map(() => '..').join('/')}${sourceSegments.length ? '/' : ''}${targetSegments.join('/')}`;
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
    && JSON.stringify(sheet.mergedRanges) === JSON.stringify(originalSheet.mergedRanges)
    && JSON.stringify(sheet.pageMargins) === JSON.stringify(originalSheet.pageMargins)
    && JSON.stringify(sheet.pageSetup) === JSON.stringify(originalSheet.pageSetup);
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
