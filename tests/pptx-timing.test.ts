import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createTimedPptxFixture } from './fixture-builders';

describe('pptx timing and transitions', () => {
  it('parses slide transition and timing node metadata', async () => {
    const presentation = parsePptx(await openPackage(createTimedPptxFixture()));
    const slide = presentation.slides[0];

    expect(slide?.transition).toEqual({ type: 'fade', speed: 'fast', advanceOnClick: true, advanceAfterMs: 7000 });
    expect(slide?.timing?.nodeCount).toBe(4);
    expect(slide?.timing?.nodes).toEqual([
      { nodeType: 'par', presetClass: 'entr', presetId: '1', id: '10', duration: '500', repeatDuration: '1500', repeatCount: '1', restart: 'always', fill: 'hold', autoReverse: true, acceleration: '10000', deceleration: '20000', triggerEvent: 'onBegin', triggerDelay: '0', triggerShapeId: '3', endTriggerEvent: 'onEnd', endTriggerDelay: '50', endTriggerShapeId: '4', targetShapeId: '2' },
      { nodeType: 'seq', concurrent: true, nextAction: 'seek', previousAction: 'none', presetClass: 'exit', presetId: '2', id: '20', duration: '750', repeatDuration: '2000', repeatCount: 'indefinite', restart: 'whenNotActive', fill: 'freeze', autoReverse: false, acceleration: '0', deceleration: '5000', triggerEvent: 'onClick', triggerDelay: '250', triggerShapeId: '3', endTriggerEvent: 'afterEffect', endTriggerDelay: '400', endTriggerShapeId: '4', targetShapeId: '2' },
      { nodeType: 'animClr', presetClass: 'emph', presetId: '3', id: '30', duration: '600', repeatDuration: '900', repeatCount: '1', restart: 'always', fill: 'hold', autoReverse: false, acceleration: '2000', deceleration: '3000', colorSpace: 'rgb', colorDirection: 'cw', triggerEvent: 'withPrevious', triggerDelay: '50', triggerShapeId: '3', endTriggerEvent: 'onEnd', endTriggerDelay: '75', endTriggerShapeId: '4', targetShapeId: '2' },
      { nodeType: 'set', presetClass: 'set', presetId: '5', id: '35', duration: '300', repeatDuration: '600', repeatCount: '1', restart: 'never', fill: 'hold', autoReverse: false, acceleration: '500', deceleration: '700', triggerEvent: 'onClick', triggerDelay: '20', triggerShapeId: '3', endTriggerEvent: 'afterEffect', endTriggerDelay: '40', endTriggerShapeId: '4', targetShapeId: '2' }
    ]);
  });

  it('renders timing and transition metadata', async () => {
    const presentation = parsePptx(await openPackage(createTimedPptxFixture()));
    const html = renderOfficeDocumentToHtml(presentation);

    expect(html).toContain('fade (fast) click:true after:7000');
    expect(html).toContain('data-timing-count="4"');
    expect(html).toContain('par:entr#10@500 repeatDur:1500×1 restart:always fill:hold autoRev:true accel:10000 decel:20000!onBegin+0^3 ~onEnd=50^4->2, seq:exit concurrent:true next:seek prev:none#20@750 repeatDur:2000×indefinite restart:whenNotActive fill:freeze autoRev:false accel:0 decel:5000!onClick+250^3 ~afterEffect=400^4->2, animClr:emph#30@600 repeatDur:900×1 restart:always fill:hold autoRev:false accel:2000 decel:3000 clr:rgb dir:cw!withPrevious+50^3 ~onEnd=75^4->2, set:set#35@300 repeatDur:600×1 restart:never fill:hold autoRev:false accel:500 decel:700!onClick+20^3 ~afterEffect=40^4->2');
  });
});
