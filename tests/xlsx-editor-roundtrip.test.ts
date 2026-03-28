import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, setWorksheetCommentText, setWorksheetTableRange } from '@ooxml/editor';
import { parseXlsx } from '@ooxml/xlsx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createCommentedXlsxFixture } from './fixture-builders';

describe('xlsx editor round-trips', () => {
  it('persists edited worksheet comments and table ranges', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createCommentedXlsxFixture())));
    setWorksheetCommentText(editor, 'Sheet1', 'B2', 'Updated comment');
    setWorksheetTableRange(editor, 'Sheet1', 'SalesTable', 'A1:B3');

    const reopened = parseXlsx(await openPackage(serializeOfficeDocument(editor.document)));
    const sheet = reopened.sheets[0];

    expect(sheet?.comments).toEqual([{ reference: 'B2', author: 'Codex', text: 'Updated comment' }]);
    expect(sheet?.tables).toEqual([{ name: 'SalesTable', range: 'A1:B3', partUri: '/xl/tables/table1.xml' }]);
  });
});
