import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createChartedXlsxFixture } from './fixture-builders';

describe('xlsx chart relationships', () => {
  it('parses worksheet chart relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createChartedXlsxFixture()));
    const chart = workbook.sheets[0]?.charts[0];

    expect(chart).toEqual({
      relationshipId: 'rIdChart1',
      drawingUri: '/xl/drawings/drawing1.xml',
      drawingNameOccurrence: 0,
      targetUri: '/xl/charts/chart1.xml',
      name: 'Sales Chart',
      chartType: 'barChart',
      title: 'Primary Chart',
      legendPosition: 'r',
      seriesNames: ['North', 'South']
    });
  });

  it('renders worksheet chart relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createChartedXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Charts: Sales Chart <barChart> [Primary Chart] legend:r {North, South} (/xl/charts/chart1.xml)');
  });
});
