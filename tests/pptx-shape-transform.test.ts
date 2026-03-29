import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createGroupedPptxFixture, createLayoutInheritedPlaceholderPptxFixture, createTransformedPptxFixture } from './fixture-builders';

describe('pptx richer shape metadata', () => {
  it('parses text and picture transforms into shape metadata', async () => {
    const presentation = parsePptx(await openPackage(createTransformedPptxFixture()));
    const slide = presentation.slides[0];
    const textShape = slide?.shapes.find((shape) => shape.name === 'Body');
    const imageShape = slide?.shapes.find((shape) => shape.media?.type === 'image');

    expect(textShape?.transform).toEqual({ x: 100, y: 200, cx: 3000, cy: 4000, rotationDeg: 30, flipV: true });
    expect(textShape?.fill).toEqual({
      kind: 'gradient',
      angleDeg: 45,
      gradientStops: [
        { position: 0, color: '#FF0000', opacity: undefined },
        { position: 100, color: '#0000FF', opacity: undefined }
      ]
    });
    expect(imageShape?.transform).toEqual({ x: 500, y: 600, cx: 7000, cy: 8000, rotationDeg: 45, flipH: true });
  });

  it('renders transform metadata into the slide projection', async () => {
    const presentation = parsePptx(await openPackage(createTransformedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('data-x="100"');
    expect(html).toContain('data-cx="3000"');
    expect(html).toContain('data-rotation-deg="30"');
    expect(html).toContain('data-flip-v="true"');
    expect(html).toContain('data-fill-gradient-stops="0:#FF0000:;100:#0000FF:"');
    expect(html).toContain('data-fill-gradient-angle="45"');
    expect(html).toContain('data-media-uri="/ppt/media/image1.png"');
    expect(html).toContain('data-x="500"');
    expect(html).toContain('data-rotation-deg="45"');
    expect(html).toContain('data-flip-h="true"');
  });

  it('parses grouped shapes into slide-level coordinates', async () => {
    const presentation = parsePptx(await openPackage(createGroupedPptxFixture()));
    const groupedShape = presentation.slides[0]?.shapes.find((shape) => shape.name === 'Grouped Title');

    expect(groupedShape?.transform).toEqual({ x: 2000, y: 3000, cx: 1600, cy: 1800, rotationDeg: 120, flipH: true });
  });

  it('inherits placeholder transform and styling from the slide layout', async () => {
    const presentation = parsePptx(await openPackage(createLayoutInheritedPlaceholderPptxFixture()));
    const inheritedShape = presentation.slides[0]?.shapes.find((shape) => shape.name === 'Body Placeholder');

    expect(presentation.slides[0]?.background?.color).toBe('#FFCC00');
    expect(inheritedShape?.transform).toEqual({ x: 1200, y: 3400, cx: 5600, cy: 1800 });
    expect(inheritedShape?.fill?.color).toBe('#123456');
    expect(inheritedShape?.textStyle?.color).toBe('#FFFFFF');
    expect(inheritedShape?.textStyle?.fontSizePt).toBe(32);
  });
});
