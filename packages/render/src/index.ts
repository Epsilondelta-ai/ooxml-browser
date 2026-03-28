import type { DocxDocument } from '@ooxml/docx';
import type { PresentationDocument } from '@ooxml/pptx';
import type { XlsxWorkbook } from '@ooxml/xlsx';

export type RenderableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export interface RenderOptions {
  activeSheetIndex?: number;
  activeSlideIndex?: number;
  showComments?: boolean;
}

export function renderOfficeDocumentToHtml(document: RenderableOfficeDocument, options: RenderOptions = {}): string {
  switch (document.kind) {
    case 'docx':
      return renderDocx(document, options);
    case 'xlsx':
      return renderWorkbook(document, options);
    case 'pptx':
      return renderPresentation(document, options);
    default:
      return '<section class="ooxml-render ooxml-render--unknown">Unsupported document</section>';
  }
}

export function mountOfficeDocument(document: RenderableOfficeDocument, target: HTMLElement, options: RenderOptions = {}): HTMLElement {
  target.innerHTML = renderOfficeDocumentToHtml(document, options);
  return target;
}

function renderDocx(document: DocxDocument, options: RenderOptions): string {
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

function renderWorkbook(workbook: XlsxWorkbook, options: RenderOptions): string {
  const sheetIndex = Math.min(options.activeSheetIndex ?? 0, Math.max(workbook.sheets.length - 1, 0));
  const sheet = workbook.sheets[sheetIndex];

  if (!sheet) {
    return '<section class="ooxml-render ooxml-render--xlsx">No sheets available</section>';
  }

  const rows = sheet.rows.map((row) => {
    const cells = row.cells.map((cell) => `<td data-ref="${escapeHtml(cell.reference)}">${escapeHtml(cell.value)}</td>`).join('');
    return `<tr><th scope="row">${row.index}</th>${cells}</tr>`;
  }).join('');

  return `<section class="ooxml-render ooxml-render--xlsx"><h2>${escapeHtml(sheet.name)}</h2><table class="ooxml-xlsx-grid"><tbody>${rows}</tbody></table></section>`;
}

function renderPresentation(presentation: PresentationDocument, options: RenderOptions): string {
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
