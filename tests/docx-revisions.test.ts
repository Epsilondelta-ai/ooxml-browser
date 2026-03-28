import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createRevisionsDocxFixture } from './fixture-builders';

describe('docx revisions', () => {
  it('parses insertion and deletion metadata from paragraphs', async () => {
    const document = parseDocx(await openPackage(createRevisionsDocxFixture()));
    const paragraph = document.stories[0]?.paragraphs[0];

    expect(paragraph?.text).toContain('Stable text');
    expect(paragraph?.text).toContain('Inserted text');
    expect(paragraph?.revisions).toEqual([
      { kind: 'insertion', id: '10', author: 'Codex', date: '2026-03-28T00:00:00Z', text: 'Inserted text' },
      { kind: 'deletion', id: '11', author: 'Codex', date: '2026-03-28T00:00:01Z', text: 'Deleted text' }
    ]);
  });

  it('renders revision markers for inserted and deleted text', async () => {
    const document = parseDocx(await openPackage(createRevisionsDocxFixture()));
    const html = renderOfficeDocumentToHtml(document);

    expect(html).toContain('data-revision-kind="insertion"');
    expect(html).toContain('[+Inserted text]');
    expect(html).toContain('data-revision-kind="deletion"');
    expect(html).toContain('[-Deleted text]');
  });
});
