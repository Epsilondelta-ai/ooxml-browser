import type { DocxDocument } from '@ooxml/docx';

import type { RenderOptions } from './types';

export function renderDocx(document: DocxDocument, options: RenderOptions): string {
  const storyMarkup = document.stories.map((story) => {
    const paragraphs = story.paragraphs
      .map((paragraph) => `<p class="ooxml-docx-paragraph">${escapeHtml(paragraph.text)}</p>`)
      .join('');
    const tables = story.tables.map((table) => {
      const rows = table.rows.map((row) => {
        const cells = row.cells.map((cell) => `<td>${escapeHtml(cell.text)}</td>`).join('');
        return `<tr>${cells}</tr>`;
      }).join('');
      return `<table class="ooxml-docx-table"><tbody>${rows}</tbody></table>`;
    }).join('');

    return `<section class="ooxml-docx-story" data-story-kind="${story.kind}">${paragraphs}${tables}</section>`;
  }).join('');

  const commentsMarkup = options.showComments === false
    ? ''
    : `<aside class="ooxml-docx-comments">${document.comments.map((comment) => `<div class="ooxml-docx-comment"><strong>${escapeHtml(comment.author ?? 'Comment')}</strong><p>${escapeHtml(comment.text)}</p></div>`).join('')}</aside>`;

  return `<article class="ooxml-render ooxml-render--docx">${storyMarkup}${commentsMarkup}</article>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
