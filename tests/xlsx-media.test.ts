import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createMediaXlsxFixture } from './fixture-builders';

describe('xlsx drawing media relationships', () => {
  it('parses worksheet media relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createMediaXlsxFixture()));
    const media = workbook.sheets[0]?.media[0];

    expect(media).toEqual({
      relationshipId: 'rIdImage1',
      drawingUri: '/xl/drawings/drawing1.xml',
      targetUri: '/xl/media/image1.png',
      type: 'image',
      name: 'Product Image'
    });
  });

  it('renders worksheet media relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createMediaXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Media: Product Image (/xl/media/image1.png)');
  });
});
