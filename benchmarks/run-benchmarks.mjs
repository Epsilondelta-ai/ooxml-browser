import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { performance } from 'node:perf_hooks';

import { zipSync } from 'fflate';
import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { parsePptx } from '@ooxml/pptx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx } from '@ooxml/xlsx';

function encodedEntries(entries) {
  const encoder = new TextEncoder();
  return Object.fromEntries(Object.entries(entries).map(([key, value]) => [key, encoder.encode(value)]));
}

function createDocxFixture() {
  return zipSync(encodedEntries({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Benchmark paragraph</w:t></w:r></w:p></w:body></w:document>`
  }));
}

function createXlsxFixture() {
  return zipSync(encodedEntries({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Bench</t></si></sst>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`
  }));
}

function createPptxFixture() {
  return zipSync(encodedEntries({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`,
    'ppt/slides/slide1.xml': `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Bench Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
  }));
}

async function benchmark(label, bytes, parse) {
  const startOpen = performance.now();
  const graph = await openPackage(bytes);
  const openMs = performance.now() - startOpen;

  const startParse = performance.now();
  const document = parse(graph);
  const parseMs = performance.now() - startParse;

  const startRender = performance.now();
  const html = renderOfficeDocumentToHtml(document);
  const renderMs = performance.now() - startRender;

  const startSerialize = performance.now();
  const serializedBytes = serializeOfficeDocument(document);
  const serializeMs = performance.now() - startSerialize;

  return {
    label,
    openMs: Number(openMs.toFixed(3)),
    parseMs: Number(parseMs.toFixed(3)),
    renderMs: Number(renderMs.toFixed(3)),
    serializeMs: Number(serializeMs.toFixed(3)),
    outputBytes: serializedBytes.byteLength,
    htmlLength: html.length
  };
}

async function representativeBenchmarks() {
  const fixtures = [
    ['docx-representative-styled', 'fixtures/docx/representative/styled.docx', parseDocx],
    ['xlsx-representative-structured', 'fixtures/xlsx/representative/structured.xlsx', parseXlsx],
    ['xlsx-representative-charted', 'fixtures/xlsx/representative/charted.xlsx', parseXlsx],
    ['pptx-representative-inherited', 'fixtures/pptx/representative/inherited.pptx', parsePptx]
  ];

  return Promise.all(fixtures.map(async ([label, fixturePath, parse]) => benchmark(label, new Uint8Array(await readFile(path.join(process.cwd(), fixturePath))), parse)));
}

const report = {
  suite: 'ooxml-benchmarks',
  generatedAt: new Date().toISOString(),
  micro: await Promise.all([
    benchmark('docx-micro', createDocxFixture(), parseDocx),
    benchmark('xlsx-micro', createXlsxFixture(), parseXlsx),
    benchmark('pptx-micro', createPptxFixture(), parsePptx)
  ]),
  representative: await representativeBenchmarks()
};

const reportsDir = path.join(process.cwd(), 'benchmarks', 'reports');
await mkdir(reportsDir, { recursive: true });
await writeFile(path.join(reportsDir, 'latest-benchmark-results.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify(report, null, 2));
