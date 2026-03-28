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
      grouping: 'clustered',
      overlap: 0,
      varyColors: true,
      gapWidth: 180,
      title: 'Primary Chart',
      legendPosition: 'r',
      categoryAxisTitle: 'Region',
      categoryAxisPosition: 'b',
      valueAxisTitle: 'Revenue',
      valueAxisPosition: 'l',
      dataLabels: { position: 'outEnd', showValue: true, showCategoryName: false },
      series: [
        { name: 'North', invertIfNegative: false, markerSymbol: 'circle', markerSize: 8 },
        { name: 'South', invertIfNegative: true, markerSymbol: 'square', markerSize: 10 }
      ],
      seriesNames: ['North', 'South']
    });
  });

  it('renders worksheet chart relationship metadata', async () => {
    const workbook = parseXlsx(await openPackage(createChartedXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Charts: Sales Chart <barChart> [Primary Chart] grouping:clustered overlap:0 varyColors:true gapWidth:180 legend:r cat:Region catPos:b val:Revenue valPos:l dLblPos:outEnd showVal:true showCat:false {North invert:false marker:circle size:8, South invert:true marker:square size:10} (/xl/charts/chart1.xml)');
  });
});
