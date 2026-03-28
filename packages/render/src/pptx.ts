import type { PresentationDocument } from '@ooxml/pptx';

import type { RenderOptions } from './types';

export function renderPresentation(presentation: PresentationDocument, options: RenderOptions): string {
  const slideIndex = Math.min(options.activeSlideIndex ?? 0, Math.max(presentation.slides.length - 1, 0));
  const slide = presentation.slides[slideIndex];

  if (!slide) {
    return '<section class="ooxml-render ooxml-render--pptx">No slides available</section>';
  }

  const shapes = slide.shapes.map((shape) => `<div class="ooxml-pptx-shape"><h3>${escapeHtml(shape.name ?? 'Shape')}</h3><p>${escapeHtml(shape.text)}</p></div>`).join('');
  const notes = slide.notesText ? `<aside class="ooxml-pptx-notes">${escapeHtml(slide.notesText)}</aside>` : '';

  return `<section class="ooxml-render ooxml-render--pptx"><header><h2>${escapeHtml(slide.title)}</h2><p>${presentation.size.cx} × ${presentation.size.cy}</p></header>${shapes}${notes}</section>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
