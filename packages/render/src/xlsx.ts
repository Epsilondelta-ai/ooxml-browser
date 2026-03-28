import { formatXlsxCellValue, type XlsxWorkbook } from '@ooxml/xlsx';

import type { RenderOptions } from './types';

export function renderWorkbook(workbook: XlsxWorkbook, options: RenderOptions): string {
  const sheetIndex = Math.min(options.activeSheetIndex ?? 0, Math.max(workbook.sheets.length - 1, 0));
  const sheet = workbook.sheets[sheetIndex];

  if (!sheet) {
    return '<section class="ooxml-render ooxml-render--xlsx">No sheets available</section>';
  }

  const frozenPaneMarkup = sheet.frozenPane ? `<p class="ooxml-xlsx-frozen-pane" data-top-left-cell="${escapeHtml(sheet.frozenPane.topLeftCell ?? '')}">Frozen pane: ${escapeHtml(sheet.frozenPane.topLeftCell ?? 'unknown')}</p>` : '';
  const mergedRangesMarkup = sheet.mergedRanges.length ? `<p class="ooxml-xlsx-merged-ranges">Merged: ${sheet.mergedRanges.map((range) => escapeHtml(range)).join(', ')}</p>` : '';

  const rows = sheet.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const displayValue = formatXlsxCellValue(workbook, cell);
      return `<td data-ref="${escapeHtml(cell.reference)}"${cell.styleIndex !== undefined ? ` data-style-index="${cell.styleIndex}"` : ''}>${escapeHtml(displayValue)}</td>`;
    }).join('');
    return `<tr><th scope="row">${row.index}</th>${cells}</tr>`;
  }).join('');

  return `<section class="ooxml-render ooxml-render--xlsx"><h2>${escapeHtml(sheet.name)}</h2>${frozenPaneMarkup}${mergedRangesMarkup}<table class="ooxml-xlsx-grid"><tbody>${rows}</tbody></table></section>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
