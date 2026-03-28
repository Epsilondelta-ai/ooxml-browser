import { findElementsByLocalName, getParsedXmlPart, relationshipById, relationshipsFor, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { WorkbookSheet, WorksheetCell, XlsxCellFormat, XlsxChart, XlsxComment, XlsxDefinedName, XlsxFrozenPane, XlsxMedia, XlsxNumberFormat, XlsxPageMargins, XlsxPageSetup, XlsxStyleTable, XlsxTable, XlsxThreadedComment, XlsxThreadedCommentPerson, XlsxWorkbook } from './model';

export function parseXlsx(graph: PackageGraph): XlsxWorkbook {
  const workbookUri = graph.rootDocumentUri ?? '/xl/workbook.xml';
  const workbookXml = getParsedXmlPart(graph, workbookUri);
  if (!workbookXml) {
    throw new Error('Workbook part is missing.');
  }

  const sharedStrings = parseSharedStrings(graph, workbookUri);
  const threadedCommentPersons = parseThreadedCommentPersons(graph, workbookUri);
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
      relationshipId ?? '',
      threadedCommentPersons
    )];
  });

  return {
    kind: 'xlsx',
    packageGraph: graph,
    sheets,
    sharedStrings,
    styles: parseStyles(graph, workbookUri),
    definedNames: parseDefinedNames(workbook as Record<string, unknown>),
    threadedCommentPersons
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

function parseSheet(graph: PackageGraph, uri: string, name: string, sheetId: number, relationshipId: string, threadedCommentPersons: XlsxThreadedCommentPerson[]): WorkbookSheet {
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
    comments: parseSheetComments(graph, uri),
    threadedComments: parseSheetThreadedComments(graph, uri, threadedCommentPersons)
  };
}

function parseThreadedCommentPersons(graph: PackageGraph, workbookUri: string): XlsxThreadedCommentPerson[] {
  const peopleRelationship = relationshipsFor(graph, workbookUri).find((relationship) => relationship.type.includes('/person'));
  if (!peopleRelationship?.resolvedTarget) {
    return [];
  }

  const xml = getParsedXmlPart(graph, peopleRelationship.resolvedTarget);
  if (!xml) {
    return [];
  }

  return findElementsByLocalName(xml.document, 'person').map((personNode) => ({
    id: xmlAttr(personNode, 'id') ?? '',
    displayName: xmlAttr(personNode, 'displayName') ?? ''
  })).filter((person) => person.id);
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
    const chartRoot = chartXml?.document;
    const chartTitleNode = chartRoot ? findElementsByLocalName(chartRoot, 'title')[0] : undefined;
    const chartTitle = chartTitleNode ? findElementsByLocalName(chartTitleNode, 't').map((node) => xmlText(node)).join('') : undefined;
    const firstSliceAngleNode = chartRoot ? findElementsByLocalName(chartRoot, 'firstSliceAng')[0] : undefined;
    const holeSizeNode = chartRoot ? findElementsByLocalName(chartRoot, 'holeSize')[0] : undefined;
    const plotVisibleOnlyNode = chartRoot ? findElementsByLocalName(chartRoot, 'plotVisOnly')[0] : undefined;
    const displayBlanksAsNode = chartRoot ? findElementsByLocalName(chartRoot, 'dispBlanksAs')[0] : undefined;
    const legendNode = chartRoot ? findElementsByLocalName(chartRoot, 'legendPos')[0] : undefined;
    const categoryAxisNode = chartRoot ? findElementsByLocalName(chartRoot, 'catAx')[0] : undefined;
    const valueAxisNode = chartRoot ? findElementsByLocalName(chartRoot, 'valAx')[0] : undefined;
    const categoryAxisTitleNode = categoryAxisNode ? findElementsByLocalName(categoryAxisNode, 'title')[0] : undefined;
    const categoryAxisPositionNode = categoryAxisNode ? findElementsByLocalName(categoryAxisNode, 'axPos')[0] : undefined;
    const valueAxisTitleNode = valueAxisNode ? findElementsByLocalName(valueAxisNode, 'title')[0] : undefined;
    const valueAxisPositionNode = valueAxisNode ? findElementsByLocalName(valueAxisNode, 'axPos')[0] : undefined;
    const chartType = chartRoot
      ? ['barChart', 'lineChart', 'pieChart', 'doughnutChart', 'areaChart', 'scatterChart']
        .find((candidate) => findElementsByLocalName(chartRoot, candidate).length > 0)
      : undefined;
    const chartTypeNode = chartType && chartRoot ? findElementsByLocalName(chartRoot, chartType)[0] : undefined;
    const smoothNode = chartTypeNode ? findElementsByLocalName(chartTypeNode, 'smooth')[0] : undefined;
    const groupingNode = chartTypeNode ? findElementsByLocalName(chartTypeNode, 'grouping')[0] : undefined;
    const overlapNode = chartTypeNode ? findElementsByLocalName(chartTypeNode, 'overlap')[0] : undefined;
    const varyColorsNode = chartTypeNode ? findElementsByLocalName(chartTypeNode, 'varyColors')[0] : undefined;
    const gapWidthNode = chartTypeNode ? findElementsByLocalName(chartTypeNode, 'gapWidth')[0] : undefined;
    const dataLabelsNode = chartTypeNode ? findElementsByLocalName(chartTypeNode, 'dLbls')[0] : undefined;
    const dataLabelPositionNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'dLblPos')[0] : undefined;
    const dataLabelSeparatorNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'separator')[0] : undefined;
    const showValueNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showVal')[0] : undefined;
    const showCategoryNameNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showCatName')[0] : undefined;
    const showSeriesNameNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showSerName')[0] : undefined;
    const showLegendKeyNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showLegendKey')[0] : undefined;
    const showLeaderLinesNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showLeaderLines')[0] : undefined;
    const showPercentNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showPercent')[0] : undefined;
    const showBubbleSizeNode = dataLabelsNode ? findElementsByLocalName(dataLabelsNode, 'showBubbleSize')[0] : undefined;
    const series = chartRoot
      ? findElementsByLocalName(chartRoot, 'ser').map((seriesNode) => {
        const textNode = findElementsByLocalName(seriesNode, 'tx')[0];
        const invertIfNegativeNode = findElementsByLocalName(seriesNode, 'invertIfNegative')[0];
        const markerNode = findElementsByLocalName(seriesNode, 'marker')[0];
        const markerSymbolNode = markerNode ? findElementsByLocalName(markerNode, 'symbol')[0] : undefined;
        const markerSizeNode = markerNode ? findElementsByLocalName(markerNode, 'size')[0] : undefined;
        return {
          name: textNode ? findElementsByLocalName(textNode, 't').map((node) => xmlText(node)).join('') : '',
          invertIfNegative: xmlAttr(invertIfNegativeNode, 'val') === '1' ? true : xmlAttr(invertIfNegativeNode, 'val') === '0' ? false : undefined,
          markerSymbol: xmlAttr(markerSymbolNode, 'val') ?? undefined,
          markerSize: (() => { const value = xmlAttr(markerSizeNode, 'val'); return value ? Number(value) : undefined; })(),
          explosion: (() => { const value = xmlAttr(findElementsByLocalName(seriesNode, 'explosion')[0], 'val'); return value ? Number(value) : undefined; })()
        };
      }).filter((seriesEntry) => Boolean(seriesEntry.name))
      : [];
    return [{
      relationshipId,
      drawingUri,
      drawingNameOccurrence,
      targetUri: target,
      name: xmlAttr(nonVisual, 'name'),
      chartType,
      smooth: xmlAttr(smoothNode, 'val') === '1' ? true : xmlAttr(smoothNode, 'val') === '0' ? false : undefined,
      grouping: xmlAttr(groupingNode, 'val') ?? undefined,
      overlap: (() => { const value = xmlAttr(overlapNode, 'val'); return value ? Number(value) : undefined; })(),
      varyColors: xmlAttr(varyColorsNode, 'val') === '1' ? true : xmlAttr(varyColorsNode, 'val') === '0' ? false : undefined,
      gapWidth: (() => { const value = xmlAttr(gapWidthNode, 'val'); return value ? Number(value) : undefined; })(),
      title: chartTitle || undefined,
      firstSliceAngle: (() => { const value = xmlAttr(firstSliceAngleNode, 'val'); return value ? Number(value) : undefined; })(),
      holeSize: (() => { const value = xmlAttr(holeSizeNode, 'val'); return value ? Number(value) : undefined; })(),
      plotVisibleOnly: xmlAttr(plotVisibleOnlyNode, 'val') === '1' ? true : xmlAttr(plotVisibleOnlyNode, 'val') === '0' ? false : undefined,
      displayBlanksAs: xmlAttr(displayBlanksAsNode, 'val') ?? undefined,
      legendPosition: xmlAttr(legendNode, 'val') ?? undefined,
      categoryAxisTitle: categoryAxisTitleNode ? findElementsByLocalName(categoryAxisTitleNode, 't').map((node) => xmlText(node)).join('') || undefined : undefined,
      categoryAxisPosition: xmlAttr(categoryAxisPositionNode, 'val') ?? undefined,
      valueAxisTitle: valueAxisTitleNode ? findElementsByLocalName(valueAxisTitleNode, 't').map((node) => xmlText(node)).join('') || undefined : undefined,
      valueAxisPosition: xmlAttr(valueAxisPositionNode, 'val') ?? undefined,
      dataLabels: dataLabelsNode ? {
        position: xmlAttr(dataLabelPositionNode, 'val') ?? undefined,
        separator: dataLabelSeparatorNode ? xmlText(dataLabelSeparatorNode) || undefined : undefined,
        showValue: xmlAttr(showValueNode, 'val') === '1' ? true : xmlAttr(showValueNode, 'val') === '0' ? false : undefined,
        showCategoryName: xmlAttr(showCategoryNameNode, 'val') === '1' ? true : xmlAttr(showCategoryNameNode, 'val') === '0' ? false : undefined,
        showSeriesName: xmlAttr(showSeriesNameNode, 'val') === '1' ? true : xmlAttr(showSeriesNameNode, 'val') === '0' ? false : undefined,
        showLegendKey: xmlAttr(showLegendKeyNode, 'val') === '1' ? true : xmlAttr(showLegendKeyNode, 'val') === '0' ? false : undefined,
        showLeaderLines: xmlAttr(showLeaderLinesNode, 'val') === '1' ? true : xmlAttr(showLeaderLinesNode, 'val') === '0' ? false : undefined,
        showPercent: xmlAttr(showPercentNode, 'val') === '1' ? true : xmlAttr(showPercentNode, 'val') === '0' ? false : undefined,
        showBubbleSize: xmlAttr(showBubbleSizeNode, 'val') === '1' ? true : xmlAttr(showBubbleSizeNode, 'val') === '0' ? false : undefined
      } : undefined,
      series,
      seriesNames: series.map((seriesEntry) => seriesEntry.name)
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

function parseSheetThreadedComments(graph: PackageGraph, sheetUri: string, threadedCommentPersons: XlsxThreadedCommentPerson[]): XlsxThreadedComment[] {
  const threadedRelationship = relationshipsFor(graph, sheetUri).find((relationship) => relationship.type.includes('/threadedComment'));
  if (!threadedRelationship?.resolvedTarget) {
    return [];
  }

  const xml = getParsedXmlPart(graph, threadedRelationship.resolvedTarget);
  if (!xml) {
    return [];
  }

  return findElementsByLocalName(xml.document, 'threadedComment').map((commentNode, index) => {
    const personId = xmlAttr(commentNode, 'personId') ?? '';
    return {
      id: xmlAttr(commentNode, 'id') ?? `threaded-${index}`,
      reference: xmlAttr(commentNode, 'ref') ?? '',
      personId,
      parentId: xmlAttr(commentNode, 'parentId') ?? undefined,
      text: findElementsByLocalName(commentNode, 'text').map((node) => xmlText(node)).join(''),
      author: threadedCommentPersons.find((person) => person.id === personId)?.displayName
    };
  });
}
