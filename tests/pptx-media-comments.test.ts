import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createEmbeddedObjectPptxFixture, createMediaPptxFixture } from './fixture-builders';

describe('pptx media and comments', () => {
  it('parses image relationships and slide comments', async () => {
    const presentation = parsePptx(await openPackage(createMediaPptxFixture()));
    const slide = presentation.slides[0];
    const mediaShape = slide?.shapes.find((shape) => shape.media?.type === 'image');

    expect(mediaShape?.media?.targetUri).toBe('/ppt/media/image1.png');
    expect(slide?.comments).toEqual([{ author: 'Codex', text: 'Review image placement', index: 0 }]);
  });

  it('renders media and comments metadata into the slide projection', async () => {
    const presentation = parsePptx(await openPackage(createMediaPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('data-media-uri="/ppt/media/image1.png"');
    expect(html).toContain('[image]');
    expect(html).toContain('data-comment-index="0"');
    expect(html).toContain('Review image placement');
  });

  it('parses embedded-object relationships into slide media metadata', async () => {
    const presentation = parsePptx(await openPackage(createEmbeddedObjectPptxFixture()));
    const embeddedShape = presentation.slides[0]?.shapes.find((shape) => shape.media?.type === 'embeddedObject');

    expect(embeddedShape?.name).toBe('Workbook Object');
    expect(embeddedShape?.media).toEqual({
      type: 'embeddedObject',
      targetUri: '/ppt/embeddings/oleObject1.bin',
      progId: 'Excel.Sheet.12'
    });
  });

  it('renders embedded-object metadata into the slide projection', async () => {
    const presentation = parsePptx(await openPackage(createEmbeddedObjectPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('data-media-type="embeddedObject"');
    expect(html).toContain('data-media-uri="/ppt/embeddings/oleObject1.bin"');
    expect(html).toContain('data-prog-id="Excel.Sheet.12"');
    expect(html).toContain('[embedded-object]');
  });
});
