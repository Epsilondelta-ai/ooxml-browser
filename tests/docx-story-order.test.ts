import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';
import { serializeOfficeDocument } from '@ooxml/serializer';

import { createDocxFixture } from './fixture-builders';

describe('docx ordered block preservation', () => {
  it('keeps paragraph/table block order in the parsed story model', async () => {
    const document = parseDocx(await openPackage(createDocxFixture()));
    const blockKinds = document.stories[0]?.blocks.map((block) => block.kind);

    expect(blockKinds).toEqual(['paragraph', 'table', 'paragraph']);
  });

  it('preserves paragraph/table order through render and serialize/reopen', async () => {
    const document = parseDocx(await openPackage(createDocxFixture()));
    const html = renderOfficeDocumentToHtml(document);
    const paragraphIndex = html.indexOf('Hello OOXML');
    const tableIndex = html.indexOf('ooxml-docx-table');
    const secondParagraphIndex = html.indexOf('Second paragraph');

    expect(paragraphIndex).toBeGreaterThanOrEqual(0);
    expect(tableIndex).toBeGreaterThan(paragraphIndex);
    expect(secondParagraphIndex).toBeGreaterThan(tableIndex);

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(document)));
    expect(reopened.stories[0]?.blocks.map((block) => block.kind)).toEqual(['paragraph', 'table', 'paragraph']);
  });
});
