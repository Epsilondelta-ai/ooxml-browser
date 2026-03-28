import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx, resolveDocxNumbering, resolveDocxStyle } from '@ooxml/docx';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createNumberedDocxFixture, createRevisionsDocxFixture, createSectionedDocxFixture, createStyledDocxFixture } from './fixture-builders';

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
