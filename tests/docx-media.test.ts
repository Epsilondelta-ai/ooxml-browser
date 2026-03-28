import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createMediaDocxFixture } from './fixture-builders';

describe('docx media relationships', () => {
  it('parses drawing images and embedded objects from story relationships', async () => {
    const document = parseDocx(await openPackage(createMediaDocxFixture()));
    const story = document.stories[0];

    expect(story?.media).toEqual([
      { relationshipId: 'rIdImage1', targetUri: '/word/media/image1.png', type: 'image', name: 'Hero Image' },
      { relationshipId: 'rIdOle1', targetUri: '/word/embeddings/oleObject1.bin', type: 'embeddedObject', progId: 'Excel.Sheet.12' }
    ]);
  });

  it('renders story media metadata', async () => {
    const document = parseDocx(await openPackage(createMediaDocxFixture()));
    const html = renderOfficeDocumentToHtml(document);

    expect(html).toContain('Media: Hero Image (/word/media/image1.png), Excel.Sheet.12 [embedded-object] (/word/embeddings/oleObject1.bin)');
  });
});
