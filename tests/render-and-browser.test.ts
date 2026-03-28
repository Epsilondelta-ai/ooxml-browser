import { describe, expect, it } from 'vitest';

import { createBrowserSession } from '@ooxml/browser';
import { inspectOfficeDocument, summarizePackageGraph } from '@ooxml/devtools';
import { parseDocx } from '@ooxml/docx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';
import { openPackage } from '@ooxml/core';

import { createDocxFixture, createPptxFixture, createXlsxFixture } from './fixture-builders';

describe('renderers', () => {
  it('renders docx content to semantic HTML', async () => {
    const graph = await openPackage(createDocxFixture());
    const document = parseDocx(graph);
    const html = renderOfficeDocumentToHtml(document);

    expect(html).toContain('Hello OOXML');
    expect(html).toContain('ooxml-docx-table');
    expect(html).toContain('Review note');
  });

  it('renders workbook and presentation summaries', async () => {
    const workbookSession = await createBrowserSession(createXlsxFixture());
    const presentationSession = await createBrowserSession(createPptxFixture());

    expect(workbookSession.renderToHtml()).toContain('Sheet1');
    expect(presentationSession.renderToHtml()).toContain('Hello Deck');
    expect(presentationSession.renderToHtml()).toContain('Speaker note');
  });
});

describe('browser session and devtools summaries', () => {
  it('creates browser sessions with summaries for docx packages', async () => {
    const session = await createBrowserSession(createDocxFixture());

    expect(session.packageSummary.officeDocumentKind).toBe('docx');
    expect(session.packageSummary.partCount).toBeGreaterThan(0);
    expect(session.documentSummary.details.paragraphs).toBe(2);
    expect(session.renderToHtml()).toContain('Hello OOXML');
  });

  it('summarizes package graphs and parsed documents directly', async () => {
    const graph = await openPackage(createDocxFixture());
    const summary = summarizePackageGraph(graph);
    const docSummary = inspectOfficeDocument(parseDocx(graph));

    expect(summary.xmlPartCount).toBeGreaterThan(0);
    expect(docSummary.primaryUnits).toBeGreaterThan(0);
  });
});
