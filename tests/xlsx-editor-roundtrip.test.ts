import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, setWorkbookCellFormula, setWorkbookDefinedNameReference, setWorkbookSheetName, setWorksheetCommentAuthor, setWorksheetCommentText, setWorksheetFrozenPane, setWorksheetMergedRanges, setWorksheetTableName, setWorksheetTableRange } from '@ooxml/editor';
import { parseXlsx } from '@ooxml/xlsx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createCommentedXlsxFixture, createStructuredXlsxFixture } from './fixture-builders';

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

  it('persists edited workbook defined-name references', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    setWorkbookDefinedNameReference(editor, 'SalesRange', 'Sheet1!$A$1:$B$9');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.definedNames[0]?.reference).toBe('Sheet1!$A$1:$B$9');
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('customWorkbookAttr="keep"');
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
