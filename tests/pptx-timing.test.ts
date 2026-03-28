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
      { nodeType: 'par', presetClass: 'entr', presetId: '1', id: '10', duration: '500', repeatCount: '1', restart: 'always', fill: 'hold', acceleration: '10000', deceleration: '20000', triggerEvent: 'onBegin', triggerDelay: '0', endTriggerEvent: 'onEnd', endTriggerDelay: '50', targetShapeId: '2' },
      { nodeType: 'seq', presetClass: 'exit', presetId: '2', id: '20', duration: '750', repeatCount: 'indefinite', restart: 'whenNotActive', fill: 'freeze', acceleration: '0', deceleration: '5000', triggerEvent: 'onClick', triggerDelay: '250', endTriggerEvent: 'afterEffect', endTriggerDelay: '400', targetShapeId: '2' }
    ]);
  });

  it('renders timing and transition metadata', async () => {
    const presentation = parsePptx(await openPackage(createTimedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('fade (fast)');
    expect(html).toContain('data-timing-count="2"');
    expect(html).toContain('par:entr#10@500×1 restart:always fill:hold accel:10000 decel:20000!onBegin+0 ~onEnd=50->2, seq:exit#20@750×indefinite restart:whenNotActive fill:freeze accel:0 decel:5000!onClick+250 ~afterEffect=400->2');
  });
});
