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
      { nodeType: 'par', presetClass: 'entr', presetId: '1', id: '10', duration: '500', repeatCount: '1', restart: 'always', fill: 'hold', triggerEvent: 'onBegin', triggerDelay: '0', targetShapeId: '2' },
      { nodeType: 'seq', presetClass: 'exit', presetId: '2', id: '20', duration: '750', repeatCount: 'indefinite', restart: 'whenNotActive', fill: 'freeze', triggerEvent: 'onClick', triggerDelay: '250', targetShapeId: '2' }
    ]);
  });

  it('renders timing and transition metadata', async () => {
    const presentation = parsePptx(await openPackage(createTimedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('fade (fast)');
    expect(html).toContain('data-timing-count="2"');
    expect(html).toContain('par:entr#10@500×1 restart:always fill:hold!onBegin+0->2, seq:exit#20@750×indefinite restart:whenNotActive fill:freeze!onClick+250->2');
  });
});
