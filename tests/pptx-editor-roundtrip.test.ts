import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { addPresentationComment, createOfficeEditor, setPresentationCommentAuthor, setPresentationCommentText, setPresentationImageTarget, setPresentationNotesText, setPresentationShapeName, setPresentationShapePlaceholderType, setPresentationShapeText, setPresentationShapeTransform, setPresentationSize, setPresentationSlideLayout, setPresentationSlideMaster, setPresentationTimingNodes, setPresentationTransition } from '@ooxml/editor';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createInheritedPptxFixture, createMediaPptxFixture, createPptxFixture, createTimedPptxFixture, createTransformedPptxFixture } from './fixture-builders';

describe('pptx editor round-trips', () => {
  it('persists edited slide comment text', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createMediaPptxFixture())));
    setPresentationCommentText(editor, 0, 0, 'Updated review');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parsePptx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    expect(reopened.slides[0]?.comments[0]?.text).toBe('Updated review');
    expect(reopened.slides[0]?.comments[0]?.author).toBe('Codex');
    expect(reopenedGraph.parts['/ppt/comments/comment1.xml']?.text).toContain('authorId="Codex"');
  });


  it('persists edited slide comment authors', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createMediaPptxFixture())));
    setPresentationCommentAuthor(editor, 0, 0, 'Reviewer');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parsePptx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    expect(reopened.slides[0]?.comments[0]?.author).toBe('Reviewer');
    expect(reopenedGraph.parts['/ppt/comments/comment1.xml']?.text).toContain('authorId="Reviewer"');
  });

  it('creates comment parts on demand for slides that start without comments', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    addPresentationComment(editor, 0, { author: 'Reviewer', text: 'Created comment' });

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parsePptx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    expect(reopened.slides[0]?.comments).toEqual([{ author: 'Reviewer', text: 'Created comment', index: 0 }]);
    expect(reopenedGraph.parts['/ppt/comments/comment1.xml']?.text).toContain('<p:text>Created comment</p:text>');
    expect(reopenedGraph.parts['/ppt/slides/_rels/slide1.xml.rels']?.text).toContain('../comments/comment1.xml');
  });

  it('patches simple slide and notes text edits without dropping root attributes', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    setPresentationShapeText(editor, 0, 0, 'Updated slide');
    setPresentationNotesText(editor, 0, 'Updated note');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parsePptx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.slides[0]?.shapes[0]?.text).toBe('Updated slide');
    expect(reopened.slides[0]?.notesText).toBe('Updated note');
    expect(reopenedGraph.parts['/ppt/slides/slide1.xml']?.text).toContain('customAttr="keep"');
    expect(reopenedGraph.parts['/ppt/notesSlides/notesSlide1.xml']?.text).toContain('customNoteAttr="keep"');
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

  it('persists edited shape placeholder metadata', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    setPresentationShapePlaceholderType(editor, 0, 0, 'subtitle');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.shapes[0]?.placeholderType).toBe('subtitle');
  });


  it('persists edited shape names', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createTransformedPptxFixture())));
    setPresentationShapeName(editor, 0, 0, 'Renamed Body');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.shapes.find((entry) => entry.name === 'Renamed Body')?.text).toBe('Transformed text');
  });

  it('persists edited image relationship targets', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createMediaPptxFixture())));
    setPresentationImageTarget(editor, 0, 1, '/ppt/media/hero2.png');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parsePptx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    expect(reopened.slides[0]?.shapes.find((entry) => entry.media?.type === 'image')?.media?.targetUri).toBe('/ppt/media/hero2.png');
    expect(reopenedGraph.parts['/ppt/slides/_rels/slide1.xml.rels']?.text).toContain('../media/hero2.png');
  });

  it('persists edited slide layout targets', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createInheritedPptxFixture())));
    setPresentationSlideLayout(editor, 0, '/ppt/slideLayouts/slideLayout2.xml');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    const reopenedGraph = await openPackage(serializeOfficeDocument(editor.document));
    expect(reopened.slides[0]?.layoutUri).toBe('/ppt/slideLayouts/slideLayout2.xml');
    expect(reopened.slides[0]?.layoutName).toBe('Two Content');
    expect(reopenedGraph.parts['/ppt/slides/_rels/slide1.xml.rels']?.text).toContain('../slideLayouts/slideLayout2.xml');
  });

  it('persists edited slide master targets', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createInheritedPptxFixture())));
    setPresentationSlideMaster(editor, 0, '/ppt/slideMasters/slideMaster2.xml');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parsePptx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);
    expect(reopened.slides[0]?.masterUri).toBe('/ppt/slideMasters/slideMaster2.xml');
    expect(reopened.slides[0]?.themeUri).toBe('/ppt/theme/theme2.xml');
    expect(reopenedGraph.parts['/ppt/slideLayouts/_rels/slideLayout1.xml.rels']?.text).toContain('../slideMasters/slideMaster2.xml');
  });

  it('persists edited slide transition metadata', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createTimedPptxFixture())));
    setPresentationTransition(editor, 0, { type: 'push', speed: 'slow' });

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.transition).toEqual({ type: 'push', speed: 'slow' });
  });


  it('persists edited presentation size metadata', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    setPresentationSize(editor, { cx: 10000000, cy: 7500000 });

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.size).toEqual({ cx: 10000000, cy: 7500000 });
  });

  it('persists edited slide timing metadata', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createTimedPptxFixture())));
    setPresentationTimingNodes(editor, 0, [
      { nodeType: 'par', presetClass: 'entr', presetId: '11' },
      { nodeType: 'seq', presetClass: 'exit', presetId: '22' },
      { nodeType: 'anim', presetClass: 'emph', presetId: '33' }
    ]);

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.timing?.nodes).toEqual([
      { nodeType: 'par', presetClass: 'entr', presetId: '11' },
      { nodeType: 'seq', presetClass: 'exit', presetId: '22' },
      { nodeType: 'anim', presetClass: 'emph', presetId: '33' }
    ]);
  });
});
