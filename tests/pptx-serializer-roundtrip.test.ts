import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createInheritedPptxFixture, createMediaPptxFixture, createTimedPptxFixture, createTransformedPptxFixture } from './fixture-builders';

describe('pptx serializer persistence', () => {
  it('preserves layout/master/theme metadata through serialize/reopen', async () => {
    const reopened = parsePptx(await openPackage(serializeOfficeDocument(parsePptx(await openPackage(createInheritedPptxFixture())))));
    const slide = reopened.slides[0];
    const theme = slide?.themeUri ? reopened.themes[slide.themeUri] : undefined;

    expect(slide?.layoutUri).toBe('/ppt/slideLayouts/slideLayout1.xml');
    expect(slide?.masterUri).toBe('/ppt/slideMasters/slideMaster1.xml');
    expect(theme?.majorLatinFont).toBe('Aptos Display');
  });

  it('preserves media and comments through serialize/reopen', async () => {
    const reopened = parsePptx(await openPackage(serializeOfficeDocument(parsePptx(await openPackage(createMediaPptxFixture())))));
    const slide = reopened.slides[0];

    expect(slide?.shapes.find((shape) => shape.media?.type === 'image')?.media?.targetUri).toBe('/ppt/media/image1.png');
    expect(slide?.comments).toEqual([{ author: 'Codex', text: 'Review image placement', index: 0 }]);
  });


  it('preserves shape transform metadata through serialize/reopen', async () => {
    const reopened = parsePptx(await openPackage(serializeOfficeDocument(parsePptx(await openPackage(createTransformedPptxFixture())))));
    const slide = reopened.slides[0];
    const textShape = slide?.shapes.find((shape) => shape.name === 'Body');
    const imageShape = slide?.shapes.find((shape) => shape.media?.type === 'image');

    expect(textShape?.transform).toEqual({ x: 100, y: 200, cx: 3000, cy: 4000, rotationDeg: 30, flipV: true });
    expect(imageShape?.transform).toEqual({ x: 500, y: 600, cx: 7000, cy: 8000, rotationDeg: 45, flipH: true });
  });

  it('preserves timing and transition metadata through serialize/reopen', async () => {
    const reopened = parsePptx(await openPackage(serializeOfficeDocument(parsePptx(await openPackage(createTimedPptxFixture())))));
    const slide = reopened.slides[0];

    expect(slide?.transition).toEqual({ type: 'fade', speed: 'fast', advanceOnClick: true, advanceAfterMs: 7000 });
    expect(slide?.timing?.nodeCount).toBe(6);
    expect(slide?.timing?.nodes.map((node) => node.nodeType)).toEqual(['par', 'seq', 'animClr', 'animMotion', 'set', 'cmd']);
  });
});
