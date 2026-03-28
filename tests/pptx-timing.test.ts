import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createTimedPptxFixture } from './fixture-builders';

describe('pptx timing and transitions', () => {
  it('parses slide transition and timing node metadata', async () => {
    const presentation = parsePptx(await openPackage(createTimedPptxFixture()));
    const slide = presentation.slides[0];

    expect(slide?.transition).toEqual({ type: 'fade', speed: 'fast' });
    expect(slide?.timing?.nodeCount).toBe(2);
    expect(slide?.timing?.nodes).toEqual([
      { nodeType: 'par', presetClass: undefined, presetId: undefined },
      { nodeType: 'seq', presetClass: undefined, presetId: undefined }
    ]);
  });

  it('renders timing and transition metadata', async () => {
    const presentation = parsePptx(await openPackage(createTimedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('data-transition-type="fade"');
    expect(html).toContain('fade (fast)');
    expect(html).toContain('data-timing-count="2"');
    expect(html).toContain('par, seq');
  });
});
