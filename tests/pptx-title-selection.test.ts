import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';

import { createPptxFixtureWithFooterUrlTitle } from './fixture-builders';

describe('pptx title selection', () => {
  it('prefers the primary title text over footer URLs and low-value footer text', async () => {
    const presentation = parsePptx(await openPackage(createPptxFixtureWithFooterUrlTitle()));

    expect(presentation.slides[0]?.title).toBe('Real Slide Title');
  });
});
