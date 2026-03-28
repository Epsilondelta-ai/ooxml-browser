import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, insertWorkbookRow, setWorkbookCellValue } from '@ooxml/editor';
import { parseXlsx } from '@ooxml/xlsx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createCommentedXlsxFixture, createStructuredXlsxFixture, createStyledXlsxFixture, createXlsxFixture } from './fixture-builders';

describe('xlsx serializer persistence', () => {
  it('persists workbook structural metadata after a row insert', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    insertWorkbookRow(editor, 'Sheet1', 2);

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseXlsx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    const sheet = reopened.sheets[0];

    expect(reopened.definedNames[0]?.reference).toBe('Sheet1!$A$1:$B$3');
    expect(sheet?.mergedRanges).toEqual(['A1:B1']);
    expect(sheet?.frozenPane?.topLeftCell).toBe('A3');
    expect(sheet?.rows[0]?.cells[1]?.formula).toBe('SUM(A1:A3)');
    expect(sheet?.rows[2]?.cells[0]?.reference).toBe('A3');
    expect(reopenedGraph.parts['/xl/workbook.xml']?.text).toContain('customWorkbookAttr="keep"');
  });

  it('preserves worksheet comments and table ranges across serialize/reopen', async () => {
    const workbook = parseXlsx(await openPackage(createCommentedXlsxFixture()));
    const reopened = parseXlsx(await openPackage(serializeOfficeDocument(workbook)));
    const sheet = reopened.sheets[0];

    expect(sheet?.comments).toEqual([{ reference: 'B2', author: 'Codex', text: 'Review this value' }]);
    expect(sheet?.tables).toEqual([{ name: 'SalesTable', range: 'A1:B2', partUri: '/xl/tables/table1.xml' }]);
  });
});

describe('xlsx shared-string preservation', () => {
  it('leaves sharedStrings.xml untouched when editing a numeric cell in the basic fixture', async () => {
    const originalBytes = createXlsxFixture();
    const originalGraph = await openPackage(originalBytes);
    const editor = createOfficeEditor(parseXlsx(originalGraph));
    setWorkbookCellValue(editor, 'Sheet1', 'B1', '99');

    const serialized = serializeOfficeDocument(editor.document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/xl/sharedStrings.xml']?.text).toBe(originalGraph.parts['/xl/sharedStrings.xml']?.text);
  });

  it('leaves sharedStrings.xml untouched when editing a numeric styled cell', async () => {
    const originalBytes = createStyledXlsxFixture();
    const originalGraph = await openPackage(originalBytes);
    const editor = createOfficeEditor(parseXlsx(originalGraph));
    setWorkbookCellValue(editor, 'Sheet1', 'B1', '99');

    const serialized = serializeOfficeDocument(editor.document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/xl/sharedStrings.xml']?.text).toBe(originalGraph.parts['/xl/sharedStrings.xml']?.text);
  });
});
