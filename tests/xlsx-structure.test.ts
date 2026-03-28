import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { extractFormulaReferences, parseXlsx, resolveDefinedName } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createStructuredXlsxFixture } from './fixture-builders';

describe('xlsx workbook structure helpers', () => {
  it('parses defined names, worksheet view metadata, and formula references', async () => {
    const workbook = parseXlsx(await openPackage(createStructuredXlsxFixture()));
    const sheet = workbook.sheets[0];

    expect(resolveDefinedName(workbook, 'SalesRange')?.reference).toBe('Sheet1!$A$1:$B$2');
    expect(sheet?.mergedRanges).toEqual(['A1:B1']);
    expect(sheet?.frozenPane).toEqual({ ySplit: 1, topLeftCell: 'A2', state: 'frozen', xSplit: undefined });
    expect(sheet?.selection).toEqual({ activeCell: 'B2', sqref: 'B2' });
    expect(sheet?.pageMargins).toEqual({ left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
    expect(sheet?.pageSetup).toEqual({ orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0, scale: undefined });
    expect(extractFormulaReferences(sheet?.rows[0]?.cells[1]?.formula ?? '')).toEqual(['A1:A2']);
  });

  it('renders merged-range, selection, and print metadata', async () => {
    const workbook = parseXlsx(await openPackage(createStructuredXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('data-top-left-cell="A2"');
    expect(html).toContain('data-active-cell="B2"');
    expect(html).toContain('Margins: left 0.5, right 0.5, top 0.75, bottom 0.75');
    expect(html).toContain('data-orientation="landscape"');
    expect(html).toContain('Merged: A1:B1');
  });
});
