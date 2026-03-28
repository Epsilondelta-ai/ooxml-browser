import type { PresentationDocument } from '@ooxml/pptx';

import type { RenderOptions } from './types';

export function renderPresentation(presentation: PresentationDocument, options: RenderOptions): string {
  const slideIndex = Math.min(options.activeSlideIndex ?? 0, Math.max(presentation.slides.length - 1, 0));
  const slide = presentation.slides[slideIndex];

  if (!slide) {
    return '<section class="ooxml-render ooxml-render--pptx">No slides available</section>';
  }

  const theme = slide.themeUri ? presentation.themes[slide.themeUri] : undefined;
  const shapeMarkup = slide.shapes.map((shape) => `<div class="ooxml-pptx-shape"${shape.placeholderType ? ` data-placeholder-type="${escapeHtml(shape.placeholderType)}"` : ''}${shape.media?.targetUri ? ` data-media-uri="${escapeHtml(shape.media.targetUri)}"` : ''}${shape.transform?.x !== undefined ? ` data-x="${shape.transform.x}"` : ''}${shape.transform?.y !== undefined ? ` data-y="${shape.transform.y}"` : ''}${shape.transform?.cx !== undefined ? ` data-cx="${shape.transform.cx}"` : ''}${shape.transform?.cy !== undefined ? ` data-cy="${shape.transform.cy}"` : ''}><h3>${escapeHtml(shape.name ?? 'Shape')}</h3><p>${escapeHtml(shape.text || (shape.media ? '[image]' : ''))}</p></div>`).join('');
  const notes = slide.notesText ? `<aside class="ooxml-pptx-notes">${escapeHtml(slide.notesText)}</aside>` : '';
  const commentsMarkup = slide.comments.length ? `<ul class="ooxml-pptx-comments">${slide.comments.map((comment) => `<li data-comment-index="${comment.index}">${escapeHtml(comment.text)}${comment.author ? ` — ${escapeHtml(comment.author)}` : ''}</li>`).join('')}</ul>` : '';
  const timingMarkup = slide.transition || slide.timing ? `<dl class="ooxml-pptx-timing">${slide.transition?.type ? `<dt>Transition</dt><dd data-transition-type="${escapeHtml(slide.transition.type)}">${escapeHtml(slide.transition.type)}${slide.transition.speed ? ` (${escapeHtml(slide.transition.speed)})` : ''}</dd>` : ''}${slide.timing ? `<dt>Timing nodes</dt><dd data-timing-count="${slide.timing.nodeCount}">${slide.timing.nodes.map((node) => `${escapeHtml(node.nodeType)}${node.presetClass ? `:${escapeHtml(node.presetClass)}` : ''}`).join(', ')}</dd>` : ''}</dl>` : '';
  const inheritanceMarkup = `<dl class="ooxml-pptx-inheritance"><dt>Layout</dt><dd>${escapeHtml(slide.layoutName ?? slide.layoutUri ?? 'none')}</dd><dt>Master</dt><dd>${escapeHtml(slide.masterName ?? slide.masterUri ?? 'none')}</dd><dt>Theme</dt><dd>${escapeHtml(theme?.name ?? theme?.colorSchemeName ?? slide.themeUri ?? 'none')}</dd>${theme?.majorLatinFont ? `<dt>Major font</dt><dd>${escapeHtml(theme.majorLatinFont)}</dd>` : ''}${theme?.minorLatinFont ? `<dt>Minor font</dt><dd>${escapeHtml(theme.minorLatinFont)}</dd>` : ''}</dl>`;

  return `<section class="ooxml-render ooxml-render--pptx"${slide.layoutUri ? ` data-layout-uri="${escapeHtml(slide.layoutUri)}"` : ''}${slide.masterUri ? ` data-master-uri="${escapeHtml(slide.masterUri)}"` : ''}${slide.themeUri ? ` data-theme-uri="${escapeHtml(slide.themeUri)}"` : ''}><header><h2>${escapeHtml(slide.title)}</h2><p>${presentation.size.cx} × ${presentation.size.cy}</p></header>${inheritanceMarkup}${timingMarkup}${shapeMarkup}${commentsMarkup}${notes}</section>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
