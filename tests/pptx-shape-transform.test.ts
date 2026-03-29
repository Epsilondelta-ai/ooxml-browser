import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createGroupedPptxFixture, createTransformedPptxFixture } from './fixture-builders';

describe('pptx richer shape metadata', () => {
  it('parses text and picture transforms into shape metadata', async () => {
    const presentation = parsePptx(await openPackage(createTransformedPptxFixture()));
    const slide = presentation.slides[0];
    const textShape = slide?.shapes.find((shape) => shape.name === 'Body');
    const imageShape = slide?.shapes.find((shape) => shape.media?.type === 'image');

    expect(textShape?.transform).toEqual({ x: 100, y: 200, cx: 3000, cy: 4000 });
    expect(imageShape?.transform).toEqual({ x: 500, y: 600, cx: 7000, cy: 8000 });
  });

  it('renders transform metadata into the slide projection', async () => {
    const presentation = parsePptx(await openPackage(createTransformedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('data-x="100"');
    expect(html).toContain('data-cx="3000"');
    expect(html).toContain('data-media-uri="/ppt/media/image1.png"');
    expect(html).toContain('data-x="500"');
  });

  it('parses grouped shapes into slide-level coordinates', async () => {
    const presentation = parsePptx(await openPackage(createGroupedPptxFixture()));
    const groupedShape = presentation.slides[0]?.shapes.find((shape) => shape.name === 'Grouped Title');

    expect(groupedShape?.transform).toEqual({ x: 2000, y: 3000, cx: 1600, cy: 1800 });
  });
});
