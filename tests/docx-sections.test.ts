import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createSectionedDocxFixture } from './fixture-builders';

describe('docx sections and headers/footers', () => {
  it('parses section properties and header/footer references', async () => {
    const document = parseDocx(await openPackage(createSectionedDocxFixture()));
    const section = document.sections[0];

    expect(section?.pageSize).toEqual({ width: 11906, height: 16838 });
    expect(section?.pageMargins).toEqual({ top: 1440, right: 1440, bottom: 1440, left: 1440 });
    expect(section?.headerReferences[0]).toEqual({ type: 'default', relationshipId: 'rIdHeader', targetUri: '/word/header1.xml' });
    expect(section?.footerReferences[0]).toEqual({ type: 'default', relationshipId: 'rIdFooter', targetUri: '/word/footer1.xml' });
  });

  it('includes header and footer stories in the render projection', async () => {
    const document = parseDocx(await openPackage(createSectionedDocxFixture()));
    const html = renderOfficeDocumentToHtml(document);

    expect(html).toContain('data-story-kind="header"');
    expect(html).toContain('Header text');
    expect(html).toContain('data-story-kind="footer"');
    expect(html).toContain('Footer text');
  });
});
