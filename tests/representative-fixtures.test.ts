import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';

import { openPackage, relationshipsFor } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { createOfficeEditor, replaceDocxParagraphText, setPresentationNotesText, setPresentationShapeText, setWorkbookCellValue, setWorksheetChartBubbleScale, setWorksheetChartTitle } from '@ooxml/editor';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx } from '@ooxml/xlsx';

async function loadFixture(path: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(path));
}

describe('representative fixture preservation', () => {
  it('preserves DOCX part graph on no-op and keeps package rels untouched on text edit', async () => {
    const originalBytes = await loadFixture('fixtures/docx/representative/basic.docx');
    const originalGraph = await openPackage(originalBytes);

    const reopenedGraph = await openPackage(serializeOfficeDocument(parseDocx(originalGraph)));
    expect(reopenedGraph.partOrder).toEqual(originalGraph.partOrder);
    expect(Object.keys(reopenedGraph.relationshipsBySource)).toEqual(Object.keys(originalGraph.relationshipsBySource));

    const editor = createOfficeEditor(parseDocx(await openPackage(originalBytes)));
    replaceDocxParagraphText(editor, 0, 0, 'Representative edit');
    const editedGraph = await openPackage(serializeOfficeDocument(editor.document));
    expect(editedGraph.parts['/_rels/.rels']?.text).toBe(originalGraph.parts['/_rels/.rels']?.text);
    expect(editedGraph.parts['/word/_rels/document.xml.rels']?.text).toBe(originalGraph.parts['/word/_rels/document.xml.rels']?.text);
  });

  it('preserves XLSX part graph on no-op and keeps workbook relationships untouched on cell edit', async () => {
    const originalBytes = await loadFixture('fixtures/xlsx/representative/basic.xlsx');
    const originalGraph = await openPackage(originalBytes);
    const reopenedGraph = await openPackage(serializeOfficeDocument(parseXlsx(originalGraph)));

    expect(reopenedGraph.partOrder).toEqual(originalGraph.partOrder);
    expect(relationshipsFor(reopenedGraph, '/xl/workbook.xml').map((rel) => rel.type)).toEqual(relationshipsFor(originalGraph, '/xl/workbook.xml').map((rel) => rel.type));

    const editor = createOfficeEditor(parseXlsx(await openPackage(originalBytes)));
    setWorkbookCellValue(editor, 'Sheet1', 'A1', 'Representative');
    const editedGraph = await openPackage(serializeOfficeDocument(editor.document));
    expect(editedGraph.parts['/xl/_rels/workbook.xml.rels']?.text).toBe(originalGraph.parts['/xl/_rels/workbook.xml.rels']?.text);
  });

  it('preserves charted XLSX relationships on no-op and keeps drawing/workbook rels untouched on chart-title edit', async () => {
    const originalBytes = await loadFixture('fixtures/xlsx/representative/charted.xlsx');
    const originalGraph = await openPackage(originalBytes);
    const reopenedGraph = await openPackage(serializeOfficeDocument(parseXlsx(originalGraph)));

    expect(reopenedGraph.partOrder).toEqual(originalGraph.partOrder);
    expect(relationshipsFor(reopenedGraph, '/xl/workbook.xml').map((rel) => rel.type)).toEqual(relationshipsFor(originalGraph, '/xl/workbook.xml').map((rel) => rel.type));
    expect(relationshipsFor(reopenedGraph, '/xl/drawings/drawing1.xml').map((rel) => rel.type)).toEqual(relationshipsFor(originalGraph, '/xl/drawings/drawing1.xml').map((rel) => rel.type));

    const editor = createOfficeEditor(parseXlsx(await openPackage(originalBytes)));
    setWorksheetChartTitle(editor, 'Sheet1', 0, 'Representative chart edit');
    const editedGraph = await openPackage(serializeOfficeDocument(editor.document));
    expect(editedGraph.parts['/xl/_rels/workbook.xml.rels']?.text).toBe(originalGraph.parts['/xl/_rels/workbook.xml.rels']?.text);
    expect(editedGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text).toBe(originalGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text);
  });

  it('preserves bubble-chart XLSX relationships on no-op and keeps drawing/workbook rels untouched on bubble-scale edit', async () => {
    const originalBytes = await loadFixture('fixtures/xlsx/representative/bubble.xlsx');
    const originalGraph = await openPackage(originalBytes);
    const reopenedGraph = await openPackage(serializeOfficeDocument(parseXlsx(originalGraph)));

    expect(reopenedGraph.partOrder).toEqual(originalGraph.partOrder);
    expect(relationshipsFor(reopenedGraph, '/xl/workbook.xml').map((rel) => rel.type)).toEqual(relationshipsFor(originalGraph, '/xl/workbook.xml').map((rel) => rel.type));
    expect(relationshipsFor(reopenedGraph, '/xl/drawings/drawing1.xml').map((rel) => rel.type)).toEqual(relationshipsFor(originalGraph, '/xl/drawings/drawing1.xml').map((rel) => rel.type));

    const editor = createOfficeEditor(parseXlsx(await openPackage(originalBytes)));
    setWorksheetChartBubbleScale(editor, 'Sheet1', 0, 180);
    const editedGraph = await openPackage(serializeOfficeDocument(editor.document));
    expect(editedGraph.parts['/xl/_rels/workbook.xml.rels']?.text).toBe(originalGraph.parts['/xl/_rels/workbook.xml.rels']?.text);
    expect(editedGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text).toBe(originalGraph.parts['/xl/drawings/_rels/drawing1.xml.rels']?.text);
  });

  it('preserves PPTX part graph on no-op and keeps slide relationships untouched on notes edit', async () => {
    const originalBytes = await loadFixture('fixtures/pptx/representative/basic.pptx');
    const originalGraph = await openPackage(originalBytes);
    const reopenedGraph = await openPackage(serializeOfficeDocument(parsePptx(originalGraph)));

    expect(reopenedGraph.partOrder).toEqual(originalGraph.partOrder);
    expect(relationshipsFor(reopenedGraph, '/ppt/slides/slide1.xml').map((rel) => rel.type)).toEqual(relationshipsFor(originalGraph, '/ppt/slides/slide1.xml').map((rel) => rel.type));

    const editor = createOfficeEditor(parsePptx(await openPackage(originalBytes)));
    setPresentationShapeText(editor, 0, 0, 'Representative slide');
    setPresentationNotesText(editor, 0, 'Representative note');
    const editedGraph = await openPackage(serializeOfficeDocument(editor.document));
    expect(editedGraph.parts['/ppt/slides/_rels/slide1.xml.rels']?.text).toBe(originalGraph.parts['/ppt/slides/_rels/slide1.xml.rels']?.text);
    expect(editedGraph.parts['/ppt/_rels/presentation.xml.rels']?.text).toBe(originalGraph.parts['/ppt/_rels/presentation.xml.rels']?.text);
  });
});
