import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx, resolveDocxNumbering } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createNumberedDocxFixture } from './fixture-builders';

describe('docx numbering graph', () => {
  it('parses numbering definitions and paragraph numbering bindings', async () => {
    const document = parseDocx(await openPackage(createNumberedDocxFixture()));
    const firstParagraph = document.stories[0]?.paragraphs[0];
    const numbering = firstParagraph ? resolveDocxNumbering(document, firstParagraph) : undefined;

    expect(firstParagraph?.numbering).toEqual({ numId: '7', level: 0 });
    expect(document.numbering.nums['7']?.abstractNumId).toBe('42');
    expect(numbering?.format).toBe('decimal');
    expect(numbering?.text).toBe('%1.');
  });

  it('renders numbering labels for numbered paragraphs', async () => {
    const document = parseDocx(await openPackage(createNumberedDocxFixture()));
    const html = renderOfficeDocumentToHtml(document);

    expect(html).toContain('data-num-id="7"');
    expect(html).toContain('1.');
    expect(html).toContain('2.');
  });
});
