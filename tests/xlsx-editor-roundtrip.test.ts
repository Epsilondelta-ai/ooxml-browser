import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, setWorkbookDefinedNameReference, setWorksheetCommentText, setWorksheetTableRange } from '@ooxml/editor';
import { parseXlsx } from '@ooxml/xlsx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createCommentedXlsxFixture, createStructuredXlsxFixture } from './fixture-builders';

describe('xlsx editor round-trips', () => {
  it('persists edited worksheet comments and table ranges', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createCommentedXlsxFixture())));
    setWorksheetCommentText(editor, 'Sheet1', 'B2', 'Updated comment');
    setWorksheetTableRange(editor, 'Sheet1', 'SalesTable', 'A1:B3');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    const sheet = reopened.sheets[0];

    expect(sheet?.comments).toEqual([{ reference: 'B2', author: 'Codex', text: 'Updated comment' }]);
    expect(sheet?.tables).toEqual([{ name: 'SalesTable', range: 'A1:B3', partUri: '/xl/tables/table1.xml' }]);
    expect(reopenedGraph.parts['/xl/comments1.xml']?.text).toContain('authorId="0"');
    expect(reopenedGraph.parts['/xl/tables/table1.xml']?.text).toContain('totalsRowShown="0"');
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
