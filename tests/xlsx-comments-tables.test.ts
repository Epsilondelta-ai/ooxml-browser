import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createCommentedXlsxFixture } from './fixture-builders';

describe('xlsx comments and tables', () => {
  it('parses worksheet comments and table relationships', async () => {
    const workbook = parseXlsx(await openPackage(createCommentedXlsxFixture()));
    const sheet = workbook.sheets[0];

    expect(sheet?.comments).toEqual([{ reference: 'B2', author: 'Codex', text: 'Review this value' }]);
    expect(sheet?.tables).toEqual([{ name: 'SalesTable', range: 'A1:B2', partUri: '/xl/tables/table1.xml' }]);
  });

  it('renders worksheet comments and table metadata', async () => {
    const workbook = parseXlsx(await openPackage(createCommentedXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Tables: SalesTable (A1:B2)');
    expect(html).toContain('data-ref="B2"');
    expect(html).toContain('Review this value');
  });
});
