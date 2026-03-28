import { formatXlsxCellValue, type XlsxWorkbook } from '@ooxml/xlsx';

import type { RenderOptions } from './types';

export function renderWorkbook(workbook: XlsxWorkbook, options: RenderOptions): string {
  const sheetIndex = Math.min(options.activeSheetIndex ?? 0, Math.max(workbook.sheets.length - 1, 0));
  const sheet = workbook.sheets[sheetIndex];

  if (!sheet) {
    return '<section class="ooxml-render ooxml-render--xlsx">No sheets available</section>';
  }

  const frozenPaneMarkup = sheet.frozenPane ? `<p class="ooxml-xlsx-frozen-pane" data-top-left-cell="${escapeHtml(sheet.frozenPane.topLeftCell ?? '')}">Frozen pane: ${escapeHtml(sheet.frozenPane.topLeftCell ?? 'unknown')}</p>` : '';
  const selectionMarkup = sheet.selection ? `<p class="ooxml-xlsx-selection" data-active-cell="${escapeHtml(sheet.selection.activeCell ?? '')}">Selection: ${escapeHtml(sheet.selection.sqref ?? sheet.selection.activeCell ?? 'unknown')}</p>` : '';
  const pageMarginsMarkup = sheet.pageMargins ? `<p class="ooxml-xlsx-page-margins">Margins: left ${escapeHtml(String(sheet.pageMargins.left ?? ''))}, right ${escapeHtml(String(sheet.pageMargins.right ?? ''))}, top ${escapeHtml(String(sheet.pageMargins.top ?? ''))}, bottom ${escapeHtml(String(sheet.pageMargins.bottom ?? ''))}</p>` : '';
  const pageSetupMarkup = sheet.pageSetup ? `<p class="ooxml-xlsx-page-setup" data-orientation="${escapeHtml(sheet.pageSetup.orientation ?? '')}">Page setup: ${escapeHtml(sheet.pageSetup.orientation ?? 'default')}${sheet.pageSetup.fitToWidth !== undefined || sheet.pageSetup.fitToHeight !== undefined ? ` (fit ${escapeHtml(String(sheet.pageSetup.fitToWidth ?? ''))}×${escapeHtml(String(sheet.pageSetup.fitToHeight ?? ''))})` : ''}</p>` : '';
  const mergedRangesMarkup = sheet.mergedRanges.length ? `<p class="ooxml-xlsx-merged-ranges">Merged: ${sheet.mergedRanges.map((range) => escapeHtml(range)).join(', ')}</p>` : '';
  const chartsMarkup = sheet.charts.length ? `<p class="ooxml-xlsx-charts">Charts: ${sheet.charts.map((chart) => `${escapeHtml(chart.name ?? chart.relationshipId)}${chart.chartType ? ` <${escapeHtml(chart.chartType)}>` : ''}${chart.title ? ` [${escapeHtml(chart.title)}]` : ''}${chart.legendPosition ? ` legend:${escapeHtml(chart.legendPosition)}` : ''}${chart.categoryAxisTitle ? ` cat:${escapeHtml(chart.categoryAxisTitle)}` : ''}${chart.valueAxisTitle ? ` val:${escapeHtml(chart.valueAxisTitle)}` : ''}${chart.seriesNames.length ? ` {${chart.seriesNames.map((seriesName) => escapeHtml(seriesName)).join(', ')}}` : ''} (${escapeHtml(chart.targetUri)})`).join(', ')}</p>` : '';
  const mediaMarkup = sheet.media.length ? `<p class="ooxml-xlsx-media">Media: ${sheet.media.map((asset) => `${escapeHtml(asset.name ?? asset.relationshipId)} (${escapeHtml(asset.targetUri)})`).join(', ')}</p>` : '';
  const tablesMarkup = sheet.tables.length ? `<p class="ooxml-xlsx-tables">Tables: ${sheet.tables.map((table) => `${escapeHtml(table.name)} (${escapeHtml(table.range)})`).join(', ')}</p>` : '';
  const commentsMarkup = sheet.comments.length ? `<ul class="ooxml-xlsx-comments">${sheet.comments.map((comment) => `<li data-ref="${escapeHtml(comment.reference)}">${escapeHtml(comment.reference)}: ${escapeHtml(comment.text)}${comment.author ? ` — ${escapeHtml(comment.author)}` : ''}</li>`).join('')}</ul>` : '';
  const threadedCommentsMarkup = sheet.threadedComments.length ? `<ul class="ooxml-xlsx-threaded-comments">${sheet.threadedComments.map((comment) => `<li data-threaded-ref="${escapeHtml(comment.reference)}">${escapeHtml(comment.reference)}: ${escapeHtml(comment.text)}${comment.author ? ` — ${escapeHtml(comment.author)}` : ''}${comment.parentId ? ` ↳ ${escapeHtml(comment.parentId)}` : ''}</li>`).join('')}</ul>` : '';

  const rows = sheet.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const displayValue = formatXlsxCellValue(workbook, cell);
      return `<td data-ref="${escapeHtml(cell.reference)}"${cell.styleIndex !== undefined ? ` data-style-index="${cell.styleIndex}"` : ''}>${escapeHtml(displayValue)}</td>`;
    }).join('');
    return `<tr><th scope="row">${row.index}</th>${cells}</tr>`;
  }).join('');

  return `<section class="ooxml-render ooxml-render--xlsx"><h2>${escapeHtml(sheet.name)}</h2>${frozenPaneMarkup}${selectionMarkup}${pageMarginsMarkup}${pageSetupMarkup}${mergedRangesMarkup}${chartsMarkup}${mediaMarkup}${tablesMarkup}${commentsMarkup}${threadedCommentsMarkup}<table class="ooxml-xlsx-grid"><tbody>${rows}</tbody></table></section>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
