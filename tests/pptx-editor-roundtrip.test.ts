import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, setPresentationCommentText, setPresentationShapeText, setPresentationShapeTransform } from '@ooxml/editor';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createMediaPptxFixture, createTransformedPptxFixture } from './fixture-builders';

describe('pptx editor round-trips', () => {
  it('persists edited slide comment text', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createMediaPptxFixture())));
    setPresentationCommentText(editor, 0, 0, 'Updated review');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.comments[0]?.text).toBe('Updated review');
  });

  it('persists edited shape text and transform metadata', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createTransformedPptxFixture())));
    setPresentationShapeText(editor, 0, 0, 'Moved shape');
    setPresentationShapeTransform(editor, 0, 0, { x: 150, y: 250, cx: 3500, cy: 4500 });

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    const shape = reopened.slides[0]?.shapes.find((entry) => entry.name === 'Body');

    expect(shape?.text).toBe('Moved shape');
    expect(shape?.transform).toEqual({ x: 150, y: 250, cx: 3500, cy: 4500 });
  });
});
