import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createInheritedPptxFixture } from './fixture-builders';

describe('pptx master/layout/theme inheritance', () => {
  it('parses slide layout/master/theme linkage and theme font metadata', async () => {
    const presentation = parsePptx(await openPackage(createInheritedPptxFixture()));
    const slide = presentation.slides[0];
    const theme = slide?.themeUri ? presentation.themes[slide.themeUri] : undefined;

    expect(slide?.layoutUri).toBe('/ppt/slideLayouts/slideLayout1.xml');
    expect(slide?.layoutName).toBe('Title Slide');
    expect(slide?.masterUri).toBe('/ppt/slideMasters/slideMaster1.xml');
    expect(slide?.themeUri).toBe('/ppt/theme/theme1.xml');
    expect(theme?.name).toBe('Office Theme');
    expect(theme?.majorLatinFont).toBe('Aptos Display');
    expect(theme?.minorLatinFont).toBe('Aptos');
    expect(slide?.shapes[0]?.placeholderType).toBe('title');
  });

  it('renders inheritance metadata and placeholder information', async () => {
    const presentation = parsePptx(await openPackage(createInheritedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('data-layout-uri="/ppt/slideLayouts/slideLayout1.xml"');
    expect(html).toContain('data-master-uri="/ppt/slideMasters/slideMaster1.xml"');
    expect(html).toContain('data-theme-uri="/ppt/theme/theme1.xml"');
    expect(html).toContain('Title Slide');
    expect(html).toContain('Aptos Display');
    expect(html).toContain('data-placeholder-type="title"');
  });
});
