import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

import { createBubbleXlsxFixture, createChartedXlsxFixture, createPieChartedXlsxFixture } from './fixture-builders';

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
      barDirection: 'col',
      plotVisibleOnly: true,
      displayBlanksAs: 'gap',
      grouping: 'clustered',
      overlap: 0,
      varyColors: true,
      gapWidth: 180,
      showDataLabelsOverMax: true,
      title: 'Primary Chart',
      legendPosition: 'r',
      categoryAxisTitle: 'Region',
      categoryAxisPosition: 'b',
      categoryAxisCrosses: 'autoZero',
      categoryAxisCrossesAt: 1,
      categoryAxisMajorGridlines: true,
      categoryAxisMinorGridlines: false,
      categoryAxisMajorTickMark: 'cross',
      categoryAxisTickMarkSkip: 2,
      categoryAxisMinorTickMark: 'none',
      categoryAxisTickLabelPosition: 'nextTo',
      categoryAxisTickLabelSkip: 2,
      categoryAxisLabelAlignment: 'ctr',
      categoryAxisNoMultiLevelLabels: true,
      categoryAxisLabelOffset: 175,
      categoryAxisDeleted: false,
      valueAxisTitle: 'Revenue',
      valueAxisPosition: 'l',
      valueAxisCrosses: 'max',
      valueAxisCrossesAt: 0,
      valueAxisCrossBetween: 'between',
      valueAxisMinimum: 0,
      valueAxisMaximum: 250,
      valueAxisMajorUnit: 50,
      valueAxisMinorUnit: 10,
      valueAxisMajorGridlines: true,
      valueAxisMinorGridlines: false,
      valueAxisMajorTickMark: 'out',
      valueAxisMinorTickMark: 'in',
      valueAxisTickLabelPosition: 'low',
      valueAxisDeleted: false,
      valueAxisDisplayUnits: 'hundreds',
      dataLabels: { position: 'outEnd', separator: ' · ', showValue: true, showCategoryName: false, showSeriesName: true, showLegendKey: false, showLeaderLines: true, showPercent: true, showBubbleSize: false },
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

    expect(html).toContain('Charts: Sales Chart <barChart> [Primary Chart] barDir:col plotVisibleOnly:true blanks:gap dLblsOverMax:true grouping:clustered overlap:0 varyColors:true gapWidth:180 legend:r cat:Region catPos:b catCross:autoZero catCrossAt:1 catMajorGrid:true catMinorGrid:false catMajorTick:cross catTickSkip:2 catMinorTick:none catLblPos:nextTo catLblSkip:2 catLblAlgn:ctr catNoMulti:true catLblOffset:175 catDelete:false val:Revenue valPos:l valCross:max valCrossAt:0 valCrossBetween:between valMin:0 valMax:250 valMajor:50 valMinor:10 valMajorGrid:true valMinorGrid:false valMajorTick:out valMinorTick:in valLblPos:low valDelete:false valDispUnits:hundreds dLblPos:outEnd dLblSep: ·  showVal:true showCat:false showSeries:true showLegendKey:false showLeaderLines:true showPercent:true showBubble:false {North invert:false marker:circle size:8, South invert:true marker:square size:10} (/xl/charts/chart1.xml)');
  });

  it('parses doughnut-specific chart metadata', async () => {
    const workbook = parseXlsx(await openPackage(createPieChartedXlsxFixture()));
    const chart = workbook.sheets[0]?.charts[0];

    expect(chart).toEqual({
      relationshipId: 'rIdChart1',
      drawingUri: '/xl/drawings/drawing1.xml',
      drawingNameOccurrence: 0,
      targetUri: '/xl/charts/chart1.xml',
      name: 'Revenue Doughnut',
      chartType: 'doughnutChart',
      varyColors: true,
      title: 'Doughnut Share',
      legendPosition: 'r',
      firstSliceAngle: 120,
      holeSize: 60,
      series: [{ name: 'Share', explosion: 25 }],
      seriesNames: ['Share']
    });
  });

  it('renders doughnut-specific chart metadata', async () => {
    const workbook = parseXlsx(await openPackage(createPieChartedXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Charts: Revenue Doughnut <doughnutChart> [Doughnut Share] firstSlice:120 holeSize:60 varyColors:true legend:r {Share explosion:25} (/xl/charts/chart1.xml)');
  });

  it('renders line-family decoration metadata', async () => {
    const workbook = parseXlsx(await openPackage(createChartedXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Alternate Chart <lineChart> [Alternate Chart] smooth:true dropLines:true hiLowLines:false serLines:true upDownBars:false plotVisibleOnly:false blanks:span grouping:standard overlap:-20 varyColors:false gapWidth:90 legend:b cat:Month catPos:l val:Bookings valPos:r dLblPos:bestFit dLblSep:/ showVal:false showCat:true showSeries:false showLegendKey:true showLeaderLines:false showPercent:false showBubble:true {Forecast invert:false marker:diamond size:6} (/xl/charts/chart2.xml)');
    expect(html).toContain('dropLines:true');
    expect(html).toContain('hiLowLines:false');
    expect(html).toContain('serLines:true');
    expect(html).toContain('upDownBars:false');
  });

  it('parses bubble-chart-specific metadata', async () => {
    const workbook = parseXlsx(await openPackage(createBubbleXlsxFixture()));
    const chart = workbook.sheets[0]?.charts[0];

    expect(chart).toEqual({
      relationshipId: 'rIdChart1',
      drawingUri: '/xl/drawings/drawing1.xml',
      drawingNameOccurrence: 0,
      targetUri: '/xl/charts/chart1.xml',
      name: 'Bubble Forecast',
      chartType: 'bubbleChart',
      bubbleScale: 140,
      showNegativeBubbles: true,
      sizeRepresents: 'area',
      varyColors: true,
      title: 'Bubble Opportunity Map',
      legendPosition: 'r',
      dataLabels: { position: 'bestFit', showValue: true, showBubbleSize: true },
      series: [{ name: 'Opportunities', invertIfNegative: false }],
      seriesNames: ['Opportunities']
    });
  });

  it('renders bubble-chart-specific metadata', async () => {
    const workbook = parseXlsx(await openPackage(createBubbleXlsxFixture()));
    const html = renderOfficeDocumentToHtml(workbook);

    expect(html).toContain('Charts: Bubble Forecast <bubbleChart> [Bubble Opportunity Map] bubbleScale:140 showNegBubbles:true sizeRep:area varyColors:true legend:r dLblPos:bestFit showVal:true showBubble:true {Opportunities invert:false} (/xl/charts/chart1.xml)');
  });
});
