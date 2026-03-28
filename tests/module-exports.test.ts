import { describe, expect, it } from 'vitest';

import { type ResolvedDrawingStyle } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { createOfficeEditor } from '@ooxml/editor';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx } from '@ooxml/xlsx';
import { openPackage } from '@ooxml/core';

import { createDocxFixture, createPptxFixture, createXlsxFixture } from './fixture-builders';

describe('module export smoke', () => {
  it('keeps re-exported package APIs working after source decomposition', async () => {
    const style: ResolvedDrawingStyle = { fillColor: { kind: 'rgb', value: '#000000' } };
    expect(style.fillColor?.value).toBe('#000000');

    const docx = parseDocx(await openPackage(createDocxFixture()));
    const xlsx = parseXlsx(await openPackage(createXlsxFixture()));
    const pptx = parsePptx(await openPackage(createPptxFixture()));

    expect(renderOfficeDocumentToHtml(docx)).toContain('Hello OOXML');
    expect(renderOfficeDocumentToHtml(xlsx)).toContain('Sheet1');
    expect(renderOfficeDocumentToHtml(pptx)).toContain('Hello Deck');

    expect(serializeOfficeDocument(createOfficeEditor(docx).document).byteLength).toBeGreaterThan(0);
  });
});
