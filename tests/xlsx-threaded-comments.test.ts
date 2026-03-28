import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createThreadedXlsxFixture } from './fixture-builders';

describe('xlsx threaded comments', () => {
  it('parses workbook persons and worksheet threaded comments', async () => {
    const workbook = parseXlsx(await openPackage(createThreadedXlsxFixture()));

    expect(workbook.threadedCommentPersons).toEqual([{ id: 'person-1', displayName: 'Avery' }]);
    expect(workbook.sheets[0]?.threadedComments).toEqual([
      { id: 'thread-1', reference: 'A1', personId: 'person-1', text: 'Discuss pipeline', author: 'Avery' }
    ]);
  });

  it('renders worksheet threaded comment metadata', async () => {
    const workbook = parseXlsx(await openPackage(createThreadedXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('data-threaded-ref="A1"');
    expect(html).toContain('Discuss pipeline');
    expect(html).toContain('Avery');
  });
});
