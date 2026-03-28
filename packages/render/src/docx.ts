import { resolveDocxNumbering, resolveDocxStyle, type DocxDocument } from '@ooxml/docx';

import type { RenderOptions } from './types';

export function renderDocx(document: DocxDocument, options: RenderOptions): string {
  const storyMarkup = document.stories.map((story) => {
    const numberingState = new Map<string, number[]>();
    const blocks = (story.blocks.length ? story.blocks : [
      ...story.paragraphs.map((paragraph) => ({ kind: 'paragraph', paragraph } as const)),
      ...story.tables.map((table) => ({ kind: 'table', table } as const))
    ]).map((block) => {
      if (block.kind === 'paragraph') {
        const paragraph = block.paragraph;
        const style = resolveDocxStyle(document, paragraph.styleId);
        const styleAttr = [
          style?.bold ? 'font-weight: 700' : '',
          style?.italic ? 'font-style: italic' : ''
        ].filter(Boolean).join('; ');

        const numbering = resolveDocxNumbering(document, paragraph);
        const label = numbering && paragraph.numbering ? renderNumberingLabel(numberingState, paragraph.numbering.numId, numbering) : '';
        const revisionsMarkup = paragraph.revisions.map((revision: typeof paragraph.revisions[number]) => `<span class="ooxml-docx-revision ooxml-docx-revision--${revision.kind}" data-revision-kind="${revision.kind}">${revision.kind === 'insertion' ? '[+' : '[-'}${escapeHtml(revision.text)}]</span>`).join(' ');
        return `<p class="ooxml-docx-paragraph"${paragraph.styleId ? ` data-style-id="${escapeHtml(paragraph.styleId)}"` : ''}${paragraph.numbering ? ` data-num-id="${escapeHtml(paragraph.numbering.numId)}" data-num-level="${paragraph.numbering.level}"` : ''}${styleAttr ? ` style="${styleAttr}"` : ''}>${label ? `<span class="ooxml-docx-numbering">${escapeHtml(label)}</span> ` : ''}${escapeHtml(paragraph.text)}${revisionsMarkup ? ` <span class="ooxml-docx-revisions">${revisionsMarkup}</span>` : ''}</p>`;
      }

      const table = block.table;
      const rows = table.rows.map((row: typeof table.rows[number]) => {
        const cells = row.cells.map((cell: typeof row.cells[number]) => `<td>${escapeHtml(cell.text)}</td>`).join('');
        return `<tr>${cells}</tr>`;
      }).join('');
      return `<table class="ooxml-docx-table"><tbody>${rows}</tbody></table>`;
    }).join('');

    const mediaMarkup = story.media.length
      ? `<p class="ooxml-docx-media">Media: ${story.media.map((asset) => `${escapeHtml(asset.type === 'embeddedObject' ? asset.progId ?? asset.relationshipId : asset.name ?? asset.relationshipId)}${asset.type === 'embeddedObject' ? ' [embedded-object]' : ''} (${escapeHtml(asset.targetUri)})`).join(', ')}</p>`
      : '';
    return `<section class="ooxml-docx-story" data-story-kind="${story.kind}">${blocks}${mediaMarkup}</section>`;
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


function renderNumberingLabel(state: Map<string, number[]>, numId: string, numbering: { level: number; format: string; text: string; start: number }): string {
  const levels = state.get(numId) ?? [];
  const levelIndex = numbering.level;
  levels[levelIndex] = (levels[levelIndex] ?? (numbering.start - 1)) + 1;
  levels.length = levelIndex + 1;
  state.set(numId, levels);

  if (numbering.format === 'bullet') {
    return numbering.text;
  }

  let label = numbering.text;
  levels.forEach((value, index) => {
    label = label.replaceAll(`%${index + 1}`, String(value));
  });
  return label;
}
