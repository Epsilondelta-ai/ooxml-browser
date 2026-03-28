import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { createOfficeEditor, setDocxCommentAuthor, setDocxCommentText } from '@ooxml/editor';
import { parseDocx, resolveDocxNumbering, resolveDocxStyle } from '@ooxml/docx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createDocxFixture, createNumberedDocxFixture, createRevisionsDocxFixture, createSectionedDocxFixture, createStyledDocxFixture } from './fixture-builders';

describe('docx serializer persistence', () => {
  it('preserves style inheritance metadata through serialize/reopen', async () => {
    const reopened = parseDocx(await openPackage(serializeOfficeDocument(parseDocx(await openPackage(createStyledDocxFixture())))));
    const style = resolveDocxStyle(reopened, 'Heading1');

    expect(style?.bold).toBe(true);
    expect(style?.italic).toBe(true);
    expect(style?.name).toBe('Heading 1');
  });

  it('preserves numbering metadata through serialize/reopen', async () => {
    const reopened = parseDocx(await openPackage(serializeOfficeDocument(parseDocx(await openPackage(createNumberedDocxFixture())))));
    const paragraph = reopened.stories[0]?.paragraphs[0];
    const numbering = paragraph ? resolveDocxNumbering(reopened, paragraph) : undefined;

    expect(paragraph?.numbering).toEqual({ numId: '7', level: 0 });
    expect(numbering?.text).toBe('%1.');
  });

  it('preserves section/header/footer metadata through serialize/reopen', async () => {
    const reopened = parseDocx(await openPackage(serializeOfficeDocument(parseDocx(await openPackage(createSectionedDocxFixture())))));
    const section = reopened.sections[0];

    expect(section?.headerReferences[0]?.targetUri).toBe('/word/header1.xml');
    expect(section?.footerReferences[0]?.targetUri).toBe('/word/footer1.xml');
    expect(reopened.stories.find((story) => story.kind === 'header')?.paragraphs[0]?.text).toBe('Header text');
  });

  it('preserves revision markers through serialize/reopen', async () => {
    const reopened = parseDocx(await openPackage(serializeOfficeDocument(parseDocx(await openPackage(createRevisionsDocxFixture())))));
    const revisions = reopened.stories[0]?.paragraphs[0]?.revisions;

    expect(revisions).toEqual([
      { kind: 'insertion', id: '10', author: 'Codex', date: '2026-03-28T00:00:00Z', text: 'Inserted text' },
      { kind: 'deletion', id: '11', author: 'Codex', date: '2026-03-28T00:00:01Z', text: 'Deleted text' }
    ]);
  });
});

describe('docx patch preservation', () => {
  it('preserves unknown document attributes when patching simple paragraph edits', async () => {
    const document = parseDocx(await openPackage(createDocxFixture()));
    document.stories[0].paragraphs[0].text = 'Patched text';
    document.stories[0].paragraphs[0].runs[0].text = 'Patched text';

    const serialized = serializeOfficeDocument(document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/word/document.xml']?.text).toContain('customAttr="keep"');
    expect(reopenedGraph.parts['/word/document.xml']?.text).toContain('Patched text');
  });

  it('leaves document.xml untouched for comment-only edits', async () => {
    const originalBytes = createDocxFixture();
    const originalGraph = await openPackage(originalBytes);
    const editor = createOfficeEditor(parseDocx(originalGraph));
    setDocxCommentText(editor, '0', 'Updated comment');

    const serialized = serializeOfficeDocument(editor.document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/word/document.xml']?.text).toBe(originalGraph.parts['/word/document.xml']?.text);
    expect(reopenedGraph.parts['/word/comments.xml']?.text).toContain('Updated comment');
  });

  it('persists edited comment authors', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createDocxFixture())));
    setDocxCommentAuthor(editor, '0', 'Reviewer');

    const serialized = serializeOfficeDocument(editor.document);
    const reopened = parseDocx(await openPackage(serialized));
    const reopenedGraph = await openPackage(serialized);

    expect(reopened.comments[0]?.author).toBe('Reviewer');
    expect(reopenedGraph.parts['/word/comments.xml']?.text).toContain('w:author="Reviewer"');
  });

  it('leaves styles.xml untouched for simple styled paragraph edits', async () => {
    const originalBytes = createStyledDocxFixture();
    const originalGraph = await openPackage(originalBytes);
    const document = parseDocx(originalGraph);
    document.stories[0].paragraphs[0].text = 'Retitled heading';
    document.stories[0].paragraphs[0].runs[0].text = 'Retitled heading';

    const serialized = serializeOfficeDocument(document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/word/styles.xml']?.text).toBe(originalGraph.parts['/word/styles.xml']?.text);
    expect(reopenedGraph.parts['/word/styles.xml']?.text).toContain('customStylesAttr="keep"');
  });

  it('leaves numbering.xml untouched for simple numbered paragraph edits', async () => {
    const originalBytes = createNumberedDocxFixture();
    const originalGraph = await openPackage(originalBytes);
    const document = parseDocx(originalGraph);
    document.stories[0].paragraphs[0].text = 'Updated list item';
    document.stories[0].paragraphs[0].runs[0].text = 'Updated list item';

    const serialized = serializeOfficeDocument(document);
    const reopenedGraph = await openPackage(serialized);

    expect(reopenedGraph.parts['/word/numbering.xml']?.text).toBe(originalGraph.parts['/word/numbering.xml']?.text);
    expect(reopenedGraph.parts['/word/numbering.xml']?.text).toContain('customNumberingAttr="keep"');
  });
});
