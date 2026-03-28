import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, removeWorkbookDefinedName, removeWorksheetComment, removeWorksheetTable, removeWorksheetThreadedComment, setWorkbookCellFormula, setWorkbookCellStyle, setWorkbookDefinedNameReference, setWorkbookDefinedNameScope, setWorkbookSheetName, setWorksheetChartBubbleScale, setWorksheetChartCategoryAxisCrosses, setWorksheetChartCategoryAxisCrossesAt, setWorksheetChartCategoryAxisDeleted, setWorksheetChartCategoryAxisLabelOffset, setWorksheetChartCategoryAxisMajorTickMark, setWorksheetChartCategoryAxisMinorTickMark, setWorksheetChartCategoryAxisPosition, setWorksheetChartCategoryAxisTickLabelPosition, setWorksheetChartCategoryAxisTitle, setWorksheetChartDataLabels, setWorksheetChartDisplayBlanksAs, setWorksheetChartFirstSliceAngle, setWorksheetChartGapWidth, setWorksheetChartGrouping, setWorksheetChartHoleSize, setWorksheetChartLegendPosition, setWorksheetChartName, setWorksheetChartOverlap, setWorksheetChartPlotVisibleOnly, setWorksheetChartScatterStyle, setWorksheetChartSeriesExplosion, setWorksheetChartSeriesInvertIfNegative, setWorksheetChartSeriesMarker, setWorksheetChartSeriesName, setWorksheetChartShowNegativeBubbles, setWorksheetChartSizeRepresents, setWorksheetChartSmooth, setWorksheetChartTarget, setWorksheetChartTitle, setWorksheetChartType, setWorksheetChartValueAxisCrossBetween, setWorksheetChartValueAxisCrosses, setWorksheetChartValueAxisCrossesAt, setWorksheetChartValueAxisDeleted, setWorksheetChartValueAxisDisplayUnits, setWorksheetChartValueAxisMajorGridlines, setWorksheetChartValueAxisMajorTickMark, setWorksheetChartValueAxisMajorUnit, setWorksheetChartValueAxisMaximum, setWorksheetChartValueAxisMinimum, setWorksheetChartValueAxisMinorGridlines, setWorksheetChartValueAxisMinorTickMark, setWorksheetChartValueAxisMinorUnit, setWorksheetChartValueAxisPosition, setWorksheetChartValueAxisTickLabelPosition, setWorksheetChartValueAxisTitle, setWorksheetChartVaryColors, setWorksheetCommentAuthor, setWorksheetCommentText, setWorksheetFrozenPane, setWorksheetMediaTarget, setWorksheetMergedRanges, setWorksheetPageMargins, setWorksheetPageSetup, setWorksheetPrintArea, setWorksheetPrintTitles, setWorksheetSelection, setWorksheetTableName, setWorksheetTableRange, setWorksheetThreadedCommentParentById, setWorksheetThreadedCommentPerson, setWorksheetThreadedCommentText, setWorksheetThreadedCommentTextById, upsertWorkbookDefinedName, upsertWorkbookThreadedCommentPerson, upsertWorksheetComment, upsertWorksheetThreadedComment } from '@ooxml/editor';
import { parseXlsx } from '@ooxml/xlsx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createBubbleXlsxFixture, createChartedXlsxFixture, createCommentedXlsxFixture, createEmbeddedObjectXlsxFixture, createMediaXlsxFixture, createPieChartedXlsxFixture, createStructuredXlsxFixture, createThreadedRepliesXlsxFixture, createThreadedXlsxFixture, createXlsxFixture } from './fixture-builders';

describe('xlsx editor round-trips', () => {
  it('persists edited worksheet comments and table ranges', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createCommentedXlsxFixture())));
    setWorksheetCommentText(editor, 'Sheet1', 'B2', 'Updated comment');
    setWorksheetTableName(editor, 'Sheet1', 'SalesTable', 'RenamedTable');
    setWorksheetTableRange(editor, 'Sheet1', 'RenamedTable', 'A1:B3');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    const sheet = reopened.sheets[0];

    expect(sheet?.comments).toEqual([{ reference: 'B2', author: 'Codex', text: 'Updated comment' }]);
    expect(sheet?.tables).toEqual([{ name: 'RenamedTable', range: 'A1:B3', partUri: '/xl/tables/table1.xml' }]);
    expect(reopenedGraph.parts['/xl/comments1.xml']?.text).toContain('authorId="0"');
    expect(reopenedGraph.parts['/xl/tables/table1.xml']?.text).toContain('displayName="RenamedTable"');
  });

  it('persists deleted worksheet tables', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createCommentedXlsxFixture())));
    removeWorksheetTable(editor, 'Sheet1', 'SalesTable');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.tables).toEqual([]);
    expect(reopenedGraph.parts['/xl/worksheets/_rels/sheet1.xml.rels']?.text).not.toContain('/relationships/table');
  });

  it('persists edited workbook defined-name references', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorkbookDefinedNameReference(editor, 'SalesRange', 'Sheet1!$A$1:$B$9');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames[0]?.reference).toBe('Sheet1!$A$1:$B$9');
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('customWorkbookAttr="keep"');
  });




  it('persists edited workbook defined-name scope metadata', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorkbookDefinedNameScope(editor, 'SalesRange', 0);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames[0]).toEqual({ name: 'SalesRange', reference: 'Sheet1!$A$1:$B$2', scopeSheetId: 0 });
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('customWorkbookAttr="keep"');
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('localSheetId="0"');
  });

  it('persists deleted workbook defined names', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    removeWorkbookDefinedName(editor, 'SalesRange');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames).toEqual([]);
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('customWorkbookAttr="keep"');
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).not.toContain('<definedName ');
  });

  it('creates workbook defined names on demand', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createXlsxFixture())));
    upsertWorkbookDefinedName(editor, 'SalesRange', 'Sheet1!$A$1:$C$5');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames).toEqual([{ name: 'SalesRange', reference: 'Sheet1!$A$1:$C$5', scopeSheetId: undefined }]);
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('<definedName name="SalesRange">Sheet1!$A$1:$C$5</definedName>');
  });

  it('persists renamed worksheets and updates dependent formula/defined-name references', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorkbookCellFormula(editor, 'Sheet1', 'B1', 'SUM(Sheet1!A1:A2)', '30');
    setWorkbookSheetName(editor, 'Sheet1', 'Summary');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.name).toBe('Summary');
    expect(reopened.definedNames[0]?.reference).toBe('Summary!$A$1:$B$2');
    expect(reopened.sheets[0]?.rows[0]?.cells[1]?.formula).toBe('SUM(Summary!A1:A2)');
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('name="Summary"');
  });

  it('persists edited worksheet comment authors', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createCommentedXlsxFixture())));
    setWorksheetCommentAuthor(editor, 'Sheet1', 'B2', 'Reviewer');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.comments).toEqual([{ reference: 'B2', author: 'Reviewer', text: 'Review this value' }]);
    expect(reopenedGraph.parts['/xl/comments1.xml']?.text).toContain('<author>Reviewer</author>');
    expect(reopenedGraph.parts['/xl/comments1.xml']?.text).toContain('authorId="0"');
  });

  it('persists deleted worksheet comments', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createCommentedXlsxFixture())));
    removeWorksheetComment(editor, 'Sheet1', 'B2');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.comments).toEqual([]);
    expect(reopenedGraph.parts['/xl/comments1.xml']?.text).not.toContain('<comment ');
  });

  it('creates worksheet comments on demand when no comments part exists', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    upsertWorksheetComment(editor, 'Sheet1', 'C3', 'Created comment', 'Reviewer');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.comments).toEqual([{ reference: 'C3', author: 'Reviewer', text: 'Created comment' }]);
    expect(reopenedGraph.parts['/xl/comments1.xml']?.text).toContain('<author>Reviewer</author>');
    expect(reopenedGraph.parts['/xl/worksheets/_rels/sheet1.xml.rels']?.text).toContain('../comments1.xml');
  });



  it('persists edited worksheet cell style indices', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorkbookCellStyle(editor, 'Sheet1', 'A1', 3);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.rows[0]?.cells[0]?.styleIndex).toBe(3);
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain(' s="3"');
  });

  it('persists edited worksheet formulas', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorkbookCellFormula(editor, 'Sheet1', 'B1', 'SUM(A1:A5)', '55');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.rows[0]?.cells[1]).toMatchObject({ reference: 'B1', formula: 'SUM(A1:A5)', value: '55' });
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('<f>SUM(A1:A5)</f>');
  });

  it('persists edited worksheet frozen pane metadata', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorksheetFrozenPane(editor, 'Sheet1', { ySplit: 2, topLeftCell: 'A3', state: 'frozen' });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.frozenPane).toEqual({ xSplit: undefined, ySplit: 2, topLeftCell: 'A3', state: 'frozen' });
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('customAttr="keep"');
  });

  it('persists worksheet print area metadata through defined names', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorksheetPrintArea(editor, 'Sheet1', '$A$1:$D$20');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames.find((entry) => entry.name === '_xlnm.Print_Area')).toEqual({
      name: '_xlnm.Print_Area',
      reference: 'Sheet1!$A$1:$D$20',
      scopeSheetId: 0
    });
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('name="_xlnm.Print_Area"');
  });

  it('persists worksheet print title metadata through defined names', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorksheetPrintTitles(editor, 'Sheet1', { rows: '$1:$2', columns: '$A:$B' });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames.find((entry) => entry.name === '_xlnm.Print_Titles')).toEqual({
      name: '_xlnm.Print_Titles',
      reference: 'Sheet1!$1:$2,Sheet1!$A:$B',
      scopeSheetId: 0
    });
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('name="_xlnm.Print_Titles"');
  });

  it('persists edited worksheet selection metadata', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorksheetSelection(editor, 'Sheet1', { activeCell: 'C3', sqref: 'C3:D4' });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.selection).toEqual({ activeCell: 'C3', sqref: 'C3:D4' });
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('activeCell="C3"');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('sqref="C3:D4"');
  });

  it('persists edited worksheet page margins and setup metadata', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorksheetPageMargins(editor, 'Sheet1', { left: 1, right: 1, top: 1.25, bottom: 1.25, header: 0.2, footer: 0.2 });
    setWorksheetPageSetup(editor, 'Sheet1', { orientation: 'portrait', paperSize: 1, fitToWidth: 2, fitToHeight: 1, scale: 95 });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.pageMargins).toEqual({ left: 1, right: 1, top: 1.25, bottom: 1.25, header: 0.2, footer: 0.2 });
    expect(reopened.sheets[0]?.pageSetup).toEqual({ orientation: 'portrait', paperSize: 1, fitToWidth: 2, fitToHeight: 1, scale: 95 });
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('orientation="portrait"');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('left="1"');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('customAttr="keep"');
  });

  it('persists edited worksheet merged ranges', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorksheetMergedRanges(editor, 'Sheet1', ['A1:B2']);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.mergedRanges).toEqual(['A1:B2']);
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('customAttr="keep"');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('mergeCell ref="A1:B2"');
  });

  it('retargets worksheet chart relationships through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartTarget(editor, 'Sheet1', 0, '/xl/charts/chart2.xml');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.targetUri).toBe('/xl/charts/chart2.xml');
    expect(reopenedGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text).toContain('../charts/chart2.xml');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('customAttr="keep"');
  });

  it('persists worksheet chart title edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartTitle(editor, 'Sheet1', 0, 'Quarterly Revenue');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.title).toBe('Quarterly Revenue');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('Quarterly Revenue');
  });

  it('persists worksheet chart frame-name edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartName(editor, 'Sheet1', 0, 'Revenue Chart');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.name).toBe('Revenue Chart');
    expect(reopenedGraph.parts['/xl/drawings/drawing1.xml']?.text).toContain('name="Revenue Chart"');
  });

  it('persists worksheet chart series-name edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartSeriesName(editor, 'Sheet1', 0, 1, 'West');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.seriesNames).toEqual(['North', 'West']);
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('West');
  });

  it('persists worksheet chart type edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartType(editor, 'Sheet1', 0, 'lineChart');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.chartType).toBe('lineChart');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('<c:lineChart>');
  });

  it('persists worksheet chart smoothing edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartType(editor, 'Sheet1', 0, 'lineChart');
    setWorksheetChartSmooth(editor, 'Sheet1', 0, true);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.chartType).toBe('lineChart');
    expect(reopened.sheets[0]?.charts[0]?.smooth).toBe(true);
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('smooth val="1"');
  });

  it('persists worksheet chart axis-crossing edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartCategoryAxisCrosses(editor, 'Sheet1', 0, 'min');
    setWorksheetChartCategoryAxisCrossesAt(editor, 'Sheet1', 0, 2);
    setWorksheetChartCategoryAxisMajorTickMark(editor, 'Sheet1', 0, 'out');
    setWorksheetChartCategoryAxisMinorTickMark(editor, 'Sheet1', 0, 'cross');
    setWorksheetChartCategoryAxisTickLabelPosition(editor, 'Sheet1', 0, 'high');
    setWorksheetChartCategoryAxisLabelOffset(editor, 'Sheet1', 0, 250);
    setWorksheetChartCategoryAxisDeleted(editor, 'Sheet1', 0, true);
    setWorksheetChartValueAxisCrosses(editor, 'Sheet1', 0, 'autoZero');
    setWorksheetChartValueAxisCrossesAt(editor, 'Sheet1', 0, 10);
    setWorksheetChartValueAxisCrossBetween(editor, 'Sheet1', 0, 'midCat');
    setWorksheetChartValueAxisMinimum(editor, 'Sheet1', 0, -5);
    setWorksheetChartValueAxisMaximum(editor, 'Sheet1', 0, 500);
    setWorksheetChartValueAxisMajorUnit(editor, 'Sheet1', 0, 25);
    setWorksheetChartValueAxisMinorUnit(editor, 'Sheet1', 0, 5);
    setWorksheetChartValueAxisMajorGridlines(editor, 'Sheet1', 0, false);
    setWorksheetChartValueAxisMinorGridlines(editor, 'Sheet1', 0, true);
    setWorksheetChartValueAxisMajorTickMark(editor, 'Sheet1', 0, 'cross');
    setWorksheetChartValueAxisMinorTickMark(editor, 'Sheet1', 0, 'out');
    setWorksheetChartValueAxisTickLabelPosition(editor, 'Sheet1', 0, 'nextTo');
    setWorksheetChartValueAxisDeleted(editor, 'Sheet1', 0, true);
    setWorksheetChartValueAxisDisplayUnits(editor, 'Sheet1', 0, 'thousands');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.categoryAxisCrosses).toBe('min');
    expect(reopened.sheets[0]?.charts[0]?.categoryAxisCrossesAt).toBe(2);
    expect(reopened.sheets[0]?.charts[0]?.categoryAxisMajorTickMark).toBe('out');
    expect(reopened.sheets[0]?.charts[0]?.categoryAxisMinorTickMark).toBe('cross');
    expect(reopened.sheets[0]?.charts[0]?.categoryAxisTickLabelPosition).toBe('high');
    expect(reopened.sheets[0]?.charts[0]?.categoryAxisLabelOffset).toBe(250);
    expect(reopened.sheets[0]?.charts[0]?.categoryAxisDeleted).toBe(true);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisCrosses).toBe('autoZero');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisCrossesAt).toBe(10);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisCrossBetween).toBe('midCat');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMinimum).toBe(-5);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMaximum).toBe(500);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMajorUnit).toBe(25);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMinorUnit).toBe(5);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMajorGridlines).toBe(false);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMinorGridlines).toBe(true);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMajorTickMark).toBe('cross');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisMinorTickMark).toBe('out');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisTickLabelPosition).toBe('nextTo');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisDeleted).toBe(true);
    expect(reopened.sheets[0]?.charts[0]?.valueAxisDisplayUnits).toBe('thousands');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('crosses val="min"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('crossesAt val="2"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('majorTickMark val="out"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('minorTickMark val="cross"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('tickLblPos val="high"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('lblOffset val="250"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('delete val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('crosses val="autoZero"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('crossesAt val="10"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('crossBetween val="midCat"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('min val="-5"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('max val="500"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('majorUnit val="25"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('minorUnit val="5"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('majorGridlines val="0"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('minorGridlines val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('majorTickMark val="cross"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('minorTickMark val="out"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('tickLblPos val="nextTo"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('dispUnits val="thousands"');
  });

  it('persists worksheet chart legend-position edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartLegendPosition(editor, 'Sheet1', 0, 't');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.legendPosition).toBe('t');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('legendPos val="t"');
  });

  it('persists worksheet chart axis-title edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartCategoryAxisTitle(editor, 'Sheet1', 0, 'Market');
    setWorksheetChartValueAxisTitle(editor, 'Sheet1', 0, 'Pipeline');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.categoryAxisTitle).toBe('Market');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisTitle).toBe('Pipeline');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('Market');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('Pipeline');
  });

  it('persists worksheet chart axis-position edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartCategoryAxisPosition(editor, 'Sheet1', 0, 'l');
    setWorksheetChartValueAxisPosition(editor, 'Sheet1', 0, 'r');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.categoryAxisPosition).toBe('l');
    expect(reopened.sheets[0]?.charts[0]?.valueAxisPosition).toBe('r');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('axPos val="l"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('axPos val="r"');
  });

  it('persists worksheet chart data-label edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartDataLabels(editor, 'Sheet1', 0, { position: 'ctr', separator: '/', showValue: false, showCategoryName: true, showSeriesName: true, showLegendKey: true, showLeaderLines: false, showPercent: true, showBubbleSize: true });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.dataLabels).toEqual({ position: 'ctr', separator: '/', showValue: false, showCategoryName: true, showSeriesName: true, showLegendKey: true, showLeaderLines: false, showPercent: true, showBubbleSize: true });
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('dLblPos val="ctr"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('<c:separator>/</c:separator>');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showVal val="0"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showCatName val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showSerName val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showLegendKey val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showLeaderLines val="0"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showPercent val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showBubbleSize val="1"');
  });

  it('persists worksheet chart vary-colors and gap-width edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartVaryColors(editor, 'Sheet1', 0, false);
    setWorksheetChartGapWidth(editor, 'Sheet1', 0, 90);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.varyColors).toBe(false);
    expect(reopened.sheets[0]?.charts[0]?.gapWidth).toBe(90);
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('varyColors val="0"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('gapWidth val="90"');
  });

  it('persists worksheet chart grouping and overlap edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartGrouping(editor, 'Sheet1', 0, 'stacked');
    setWorksheetChartOverlap(editor, 'Sheet1', 0, 100);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.grouping).toBe('stacked');
    expect(reopened.sheets[0]?.charts[0]?.overlap).toBe(100);
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('grouping val="stacked"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('overlap val="100"');
  });

  it('persists worksheet chart plot-visible and blanks-as edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartPlotVisibleOnly(editor, 'Sheet1', 0, false);
    setWorksheetChartDisplayBlanksAs(editor, 'Sheet1', 0, 'zero');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.plotVisibleOnly).toBe(false);
    expect(reopened.sheets[0]?.charts[0]?.displayBlanksAs).toBe('zero');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('plotVisOnly val="0"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('dispBlanksAs val="zero"');
  });

  it('persists worksheet chart per-series invert-if-negative edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartSeriesInvertIfNegative(editor, 'Sheet1', 0, 0, true);
    setWorksheetChartSeriesInvertIfNegative(editor, 'Sheet1', 0, 1, false);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.series).toEqual([
      { name: 'North', invertIfNegative: true, markerSymbol: 'circle', markerSize: 8 },
      { name: 'South', invertIfNegative: false, markerSymbol: 'square', markerSize: 10 }
    ]);
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('invertIfNegative val="1"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('invertIfNegative val="0"');
  });

  it('persists worksheet chart series-marker edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createChartedXlsxFixture())));
    setWorksheetChartSeriesMarker(editor, 'Sheet1', 0, 0, { symbol: 'triangle', size: 12 });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.series[0]).toEqual({ name: 'North', invertIfNegative: false, markerSymbol: 'triangle', markerSize: 12 });
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('symbol val="triangle"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('size val="12"');
  });

  it('persists doughnut first-slice, hole-size, and explosion edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createPieChartedXlsxFixture())));
    setWorksheetChartFirstSliceAngle(editor, 'Sheet1', 0, 180);
    setWorksheetChartHoleSize(editor, 'Sheet1', 0, 72);
    setWorksheetChartSeriesExplosion(editor, 'Sheet1', 0, 0, 40);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.firstSliceAngle).toBe(180);
    expect(reopened.sheets[0]?.charts[0]?.holeSize).toBe(72);
    expect(reopened.sheets[0]?.charts[0]?.series).toEqual([{ name: 'Share', explosion: 40 }]);
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('firstSliceAng val="180"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('holeSize val="72"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('explosion val="40"');
  });

  it('persists bubble-chart policy edits through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createBubbleXlsxFixture())));
    setWorksheetChartScatterStyle(editor, 'Sheet1', 0, 'lineMarker');
    setWorksheetChartBubbleScale(editor, 'Sheet1', 0, 180);
    setWorksheetChartShowNegativeBubbles(editor, 'Sheet1', 0, false);
    setWorksheetChartSizeRepresents(editor, 'Sheet1', 0, 'w');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.charts[0]?.scatterStyle).toBe('lineMarker');
    expect(reopened.sheets[0]?.charts[0]?.bubbleScale).toBe(180);
    expect(reopened.sheets[0]?.charts[0]?.showNegativeBubbles).toBe(false);
    expect(reopened.sheets[0]?.charts[0]?.sizeRepresents).toBe('w');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('scatterStyle val="lineMarker"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('bubbleScale val="180"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('showNegBubbles val="0"');
    expect(reopenedGraph.parts['/xl/charts/chart1.xml']?.text).toContain('sizeRepresents val="w"');
  });

  it('retargets worksheet media relationships through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createMediaXlsxFixture())));
    setWorksheetMediaTarget(editor, 'Sheet1', 0, '/xl/media/image2.png');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.media[0]?.targetUri).toBe('/xl/media/image2.png');
    expect(reopenedGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text).toContain('../media/image2.png');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('customAttr="keep"');
  });

  it('retargets worksheet embedded-object relationships through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createEmbeddedObjectXlsxFixture())));
    setWorksheetMediaTarget(editor, 'Sheet1', 0, '/xl/embeddings/oleObject2.bin');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.media[0]?.targetUri).toBe('/xl/embeddings/oleObject2.bin');
    expect(reopened.sheets[0]?.media[0]?.type).toBe('embeddedObject');
    expect(reopened.sheets[0]?.media[0]?.progId).toBe('Excel.Sheet.12');
    expect(reopenedGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text).toContain('../embeddings/oleObject2.bin');
  });

  it('persists threaded comment text and person metadata through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createThreadedXlsxFixture())));
    upsertWorkbookThreadedCommentPerson(editor, 'person-2', 'Blair');
    setWorksheetThreadedCommentText(editor, 'Sheet1', 'A1', 'Updated threaded note');
    setWorksheetThreadedCommentPerson(editor, 'Sheet1', 'A1', 'person-2');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.threadedCommentPersons).toEqual([
      { id: 'person-1', displayName: 'Avery' },
      { id: 'person-2', displayName: 'Blair' }
    ]);
    expect(reopened.sheets[0]?.threadedComments).toEqual([
      { id: 'thread-1', reference: 'A1', personId: 'person-2', text: 'Updated threaded note', author: 'Blair' }
    ]);
    expect(reopenedGraph.parts['/xl/persons/person1.xml']?.text).toContain('displayName="Blair"');
    expect(reopenedGraph.parts['/xl/threadedComments/threadedComment1.xml']?.text).toContain('Updated threaded note');
  });

  it('creates and removes worksheet threaded comments through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createXlsxFixture())));
    upsertWorkbookThreadedCommentPerson(editor, 'person-1', 'Avery');
    upsertWorksheetThreadedComment(editor, 'Sheet1', 'A1', 'Brand new thread', 'person-1');
    removeWorksheetThreadedComment(editor, 'Sheet1', 'A1');
    upsertWorksheetThreadedComment(editor, 'Sheet1', 'B2', 'Escalate follow-up', 'person-1');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.threadedCommentPersons).toEqual([{ id: 'person-1', displayName: 'Avery' }]);
    expect(reopened.sheets[0]?.threadedComments).toEqual([
      { id: 'threaded-b2', reference: 'B2', personId: 'person-1', text: 'Escalate follow-up', author: 'Avery' }
    ]);
    expect(reopenedGraph.parts['/xl/worksheets/_rels/sheet1.xml.rels']?.text).toContain('/threadedComment');
    expect(reopenedGraph.parts['/xl/threadedComments/threadedComment1.xml']?.text).toContain('Escalate follow-up');
  });

  it('persists threaded reply-parent metadata through save flows', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createThreadedRepliesXlsxFixture())));
    setWorksheetThreadedCommentParentById(editor, 'Sheet1', 'thread-2', 'thread-1');
    setWorksheetThreadedCommentTextById(editor, 'Sheet1', 'thread-2', 'Need legal review');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.sheets[0]?.threadedComments).toEqual([
      { id: 'thread-1', reference: 'A1', personId: 'person-1', text: 'Discuss pipeline', author: 'Avery' },
      { id: 'thread-2', reference: 'A1', personId: 'person-1', parentId: 'thread-1', text: 'Need legal review', author: 'Avery' }
    ]);
    expect(reopenedGraph.parts['/xl/threadedComments/threadedComment1.xml']?.text).toContain('parentId="thread-1"');
  });
});

describe('xlsx worksheet patch preservation', () => {
  it('preserves unknown worksheet attributes when patching simple cell edits', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    editor.document.sheets[0].rows[0].cells[0].value = '15';
    editor.document.sheets[0].rows[0].cells[0].type = 'n';

    const serialized = serializeOfficeDocument(editor.document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('customAttr="keep"');
    expect(reopenedGraph.parts['/xl/worksheets/sheet1.xml']?.text).toContain('<v>15</v>');
  });
});
