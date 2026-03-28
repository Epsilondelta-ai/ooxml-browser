import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { formatXlsxCellValue, parseXlsx, resolveXlsxCellFormat } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createStyledXlsxFixture } from './fixture-builders';

describe('xlsx style table and number formats', () => {
  it('parses number formats and resolves cell formats', async () => {
    const workbook = parseXlsx(await openPackage(createStyledXlsxFixture()));
    const styledCell = workbook.sheets[0]?.rows[0]?.cells[1];
    const style = styledCell ? resolveXlsxCellFormat(workbook, styledCell) : undefined;

    expect(workbook.styles.numberFormats[164]?.code).toBe('0.00%');
    expect(style?.numFmtId).toBe(164);
    expect(formatXlsxCellValue(workbook, styledCell!)).toBe('25.00%');
  });

  it('renders styled numeric display values with style metadata', async () => {
    const workbook = parseXlsx(await openPackage(createStyledXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('data-style-index="1"');
    expect(html).toContain('25.00%');
  });
});
