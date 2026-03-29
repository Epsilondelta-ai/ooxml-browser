import { describe, expect, it } from 'vitest';

import { openPackage, relationshipsFor } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { parsePptx } from '@ooxml/pptx';
import { parseXlsx } from '@ooxml/xlsx';

import { createDocxFixture, createPptxFixture, createPptxFixtureWithLeadingExtendedPropertiesRelationship, createXlsxFixture } from './fixture-builders';

describe('OPC package graph', () => {
  it('opens a docx package and resolves relationships', async () => {
    const graph = await openPackage(createDocxFixture());

    expect(graph.officeDocumentKind).toBe('docx');
    expect(graph.rootDocumentUri).toBe('/word/document.xml');
    expect(relationshipsFor(graph, 'package')[0]?.resolvedTarget).toBe('/word/document.xml');
    expect(graph.parts['/word/comments.xml']?.contentType).toContain('comments');
  });

  it('opens a workbook package and discovers shared strings', async () => {
    const graph = await openPackage(createXlsxFixture());

    expect(graph.officeDocumentKind).toBe('xlsx');
    expect(graph.rootDocumentUri).toBe('/xl/workbook.xml');
    expect(graph.parts['/xl/sharedStrings.xml']).toBeDefined();
  });

  it('opens a presentation package and resolves notes relationships', async () => {
    const graph = await openPackage(createPptxFixture());

    expect(graph.officeDocumentKind).toBe('pptx');
    expect(graph.rootDocumentUri).toBe('/ppt/presentation.xml');
    expect(graph.parts['/ppt/notesSlides/notesSlide1.xml']).toBeDefined();
  });

  it('uses the exact officeDocument relationship instead of matching extended-properties first', async () => {
    const graph = await openPackage(createPptxFixtureWithLeadingExtendedPropertiesRelationship());

    expect(graph.officeDocumentKind).toBe('pptx');
    expect(graph.rootDocumentUri).toBe('/ppt/presentation.xml');
  });
});

describe('format parsers', () => {
  it('parses docx stories, tables, and comments', async () => {
    const graph = await openPackage(createDocxFixture());
    const document = parseDocx(graph);

    expect(document.stories[0]?.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Hello OOXML', 'Second paragraph']);
    expect(document.stories[0]?.tables[0]?.rows[0]?.cells.map((cell) => cell.text)).toEqual(['Cell 1', 'Cell 2']);
    expect(document.comments[0]?.text).toBe('Review note');
  });

  it('parses xlsx sheets, shared strings, and formulas', async () => {
    const graph = await openPackage(createXlsxFixture());
    const workbook = parseXlsx(graph);

    expect(workbook.sharedStrings).toEqual(['Hello Sheet']);
    expect(workbook.sheets[0]?.rows[0]?.cells.map((cell) => cell.value)).toEqual(['Hello Sheet', '42', '50']);
    expect(workbook.sheets[0]?.rows[0]?.cells[2]?.formula).toBe('SUM(B1,8)');
  });

  it('parses pptx slides and speaker notes', async () => {
    const graph = await openPackage(createPptxFixture());
    const presentation = parsePptx(graph);

    expect(presentation.size).toEqual({ cx: 9144000, cy: 6858000 });
    expect(presentation.slides[0]?.title).toBe('Hello Deck');
    expect(presentation.slides[0]?.notesText).toBe('Speaker note');
  });
});
