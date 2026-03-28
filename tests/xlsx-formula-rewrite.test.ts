import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, insertWorkbookRow } from '@ooxml/editor';
import { parseXlsx } from '@ooxml/xlsx';

import { createStructuredXlsxFixture } from './fixture-builders';

describe('xlsx formula/reference rewrite helpers', () => {
  it('shifts formula references, row indices, merged ranges, frozen panes, and defined names when inserting a row', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createStructuredXlsxFixture())));
    insertWorkbookRow(editor, 'Sheet1', 2);

    const workbook = editor.document;
    const sheet = workbook.sheets[0];

    expect(sheet?.rows.map((row) => row.index)).toEqual([1, 2, 3]);
    expect(sheet?.rows[0]?.cells[1]?.formula).toBe('SUM(A1:A3)');
    expect(sheet?.rows[1]?.cells).toEqual([]);
    expect(sheet?.rows[2]?.cells[0]?.reference).toBe('A3');
    expect(sheet?.mergedRanges).toEqual(['A1:B1']);
    expect(sheet?.frozenPane?.topLeftCell).toBe('A3');
    expect(workbook.definedNames[0]?.reference).toBe('Sheet1!$A$1:$B$3');
  });
});
