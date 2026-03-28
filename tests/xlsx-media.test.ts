import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createEmbeddedObjectXlsxFixture, createMediaXlsxFixture } from './fixture-builders';

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

  it('parses worksheet embedded-object relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createEmbeddedObjectXlsxFixture()));
    const media = workbook.sheets[0]?.media[0];

    expect(media).toEqual({
      relationshipId: 'rIdOle',
      drawingUri: '/xl/drawings/drawing1.xml',
      targetUri: '/xl/embeddings/oleObject1.bin',
      type: 'embeddedObject',
      name: 'Workbook Object',
      progId: 'Excel.Sheet.12'
    });
  });

  it('renders worksheet embedded-object relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createEmbeddedObjectXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Workbook Object [embedded-object] progId:Excel.Sheet.12 (/xl/embeddings/oleObject1.bin)');
  });
});
