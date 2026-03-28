import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx, resolveDocxStyle } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createStyledDocxFixture } from './fixture-builders';

describe('docx style graph', () => {
  it('parses style inheritance and resolves merged style flags', async () => {
    const document = parseDocx(await openPackage(createStyledDocxFixture()));
    const style = resolveDocxStyle(document, 'Heading1');

    expect(document.styles.Base?.bold).toBe(true);
    expect(style?.bold).toBe(true);
    expect(style?.italic).toBe(true);
    expect(style?.name).toBe('Heading 1');
  });

  it('renders paragraph style metadata and inherited style flags', async () => {
    const document = parseDocx(await openPackage(createStyledDocxFixture()));
    const html = renderOfficeDocumentToHtml(document);

    expect(html).toContain('data-style-id="Heading1"');
    expect(html).toContain('font-weight: 700');
    expect(html).toContain('font-style: italic');
  });
});
