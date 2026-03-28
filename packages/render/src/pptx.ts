import type { PresentationDocument } from '@ooxml/pptx';

import type { RenderOptions } from './types';

export function renderPresentation(presentation: PresentationDocument, options: RenderOptions): string {
  const slideIndex = Math.min(options.activeSlideIndex ?? 0, Math.max(presentation.slides.length - 1, 0));
  const slide = presentation.slides[slideIndex];

  if (!slide) {
    return '<section class="ooxml-render ooxml-render--pptx">No slides available</section>';
  }

  const theme = slide.themeUri ? presentation.themes[slide.themeUri] : undefined;
  const shapeMarkup = slide.shapes.map((shape) => `<div class="ooxml-pptx-shape"${shape.placeholderType ? ` data-placeholder-type="${escapeHtml(shape.placeholderType)}"` : ''}><h3>${escapeHtml(shape.name ?? 'Shape')}</h3><p>${escapeHtml(shape.text)}</p></div>`).join('');
  const notes = slide.notesText ? `<aside class="ooxml-pptx-notes">${escapeHtml(slide.notesText)}</aside>` : '';
  const inheritanceMarkup = `<dl class="ooxml-pptx-inheritance"><dt>Layout</dt><dd>${escapeHtml(slide.layoutName ?? slide.layoutUri ?? 'none')}</dd><dt>Master</dt><dd>${escapeHtml(slide.masterName ?? slide.masterUri ?? 'none')}</dd><dt>Theme</dt><dd>${escapeHtml(theme?.name ?? theme?.colorSchemeName ?? slide.themeUri ?? 'none')}</dd>${theme?.majorLatinFont ? `<dt>Major font</dt><dd>${escapeHtml(theme.majorLatinFont)}</dd>` : ''}${theme?.minorLatinFont ? `<dt>Minor font</dt><dd>${escapeHtml(theme.minorLatinFont)}</dd>` : ''}</dl>`;

  return `<section class="ooxml-render ooxml-render--pptx"${slide.layoutUri ? ` data-layout-uri="${escapeHtml(slide.layoutUri)}"` : ''}${slide.masterUri ? ` data-master-uri="${escapeHtml(slide.masterUri)}"` : ''}${slide.themeUri ? ` data-theme-uri="${escapeHtml(slide.themeUri)}"` : ''}><header><h2>${escapeHtml(slide.title)}</h2><p>${presentation.size.cx} × ${presentation.size.cy}</p></header>${inheritanceMarkup}${shapeMarkup}${notes}</section>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
