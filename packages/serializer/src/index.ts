import { clonePackageGraph, serializePackageGraph, updatePackagePartText } from '@ooxml/core';
import type { DocxComment, DocxDocument, DocxParagraph, DocxTable } from '@ooxml/docx';
import type { PresentationDocument, PresentationSlide, SlideShape } from '@ooxml/pptx';
import type { WorkbookSheet, WorksheetCell, XlsxWorkbook } from '@ooxml/xlsx';

export type SerializableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export function serializeOfficeDocument(document: SerializableOfficeDocument): Uint8Array {
  switch (document.kind) {
    case 'docx':
      return serializeDocx(document);
    case 'xlsx':
      return serializeXlsx(document);
    case 'pptx':
      return serializePptx(document);
  }
}

function serializeDocx(document: DocxDocument): Uint8Array {
  const graph = clonePackageGraph(document.packageGraph);
  const mainStory = document.stories.find((story) => story.kind === 'document');

  if (mainStory) {
    updatePackagePartText(
      graph,
      mainStory.uri,
      buildDocxStoryXml(mainStory.paragraphs, mainStory.tables),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
    );
  }

  const commentsUri = '/word/comments.xml';
  if (graph.parts[commentsUri]) {
    updatePackagePartText(
      graph,
      commentsUri,
      buildCommentsXml(document.comments),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    );
  }

  return serializePackageGraph(graph);
}

function serializeXlsx(workbook: XlsxWorkbook): Uint8Array {
  const graph = clonePackageGraph(workbook.packageGraph);
  const sharedStringPool = createSharedStringPool(workbook);
  const sharedStringsUri = '/xl/sharedStrings.xml';

  const hasSharedStringsPart = Boolean(graph.parts[sharedStringsUri]);

  if (hasSharedStringsPart) {
    updatePackagePartText(
      graph,
      sharedStringsUri,
      buildSharedStringsXml(sharedStringPool.values),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
    );
  }

  for (const sheet of workbook.sheets) {
    updatePackagePartText(
      graph,
      sheet.uri,
      buildWorksheetXml(sheet, sharedStringPool.indexByValue, hasSharedStringsPart),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
    );
  }

  return serializePackageGraph(graph);
}

function serializePptx(presentation: PresentationDocument): Uint8Array {
  const graph = clonePackageGraph(presentation.packageGraph);

  for (const slide of presentation.slides) {
    updatePackagePartText(
      graph,
      slide.uri,
      buildSlideXml(slide),
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    );

    const notesUri = slide.notesUri;
    if (notesUri && graph.parts[notesUri]) {
      updatePackagePartText(
        graph,
        notesUri,
        buildNotesXml(slide.notesText),
        'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
      );
    }
  }

  return serializePackageGraph(graph);
}

function buildDocxStoryXml(paragraphs: DocxParagraph[], tables: DocxTable[]): string {
  const paragraphXml = paragraphs.map((paragraph) => `<w:p>${paragraph.styleId ? `<w:pPr><w:pStyle w:val="${escapeXml(paragraph.styleId)}"/></w:pPr>` : ''}${paragraph.runs.map((run) => `<w:r>${run.bold || run.italic ? `<w:rPr>${run.bold ? '<w:b/>' : ''}${run.italic ? '<w:i/>' : ''}</w:rPr>` : ''}<w:t>${escapeXml(run.text)}</w:t></w:r>`).join('')}</w:p>`).join('');
  const tableXml = tables.map((table) => `<w:tbl>${table.rows.map((row) => `<w:tr>${row.cells.map((cell) => `<w:tc><w:p><w:r><w:t>${escapeXml(cell.text)}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphXml}${tableXml}</w:body></w:document>`;
}

function buildCommentsXml(comments: DocxComment[]): string {
  const commentXml = comments.map((comment) => `<w:comment w:id="${escapeXml(comment.id)}"${comment.author ? ` w:author="${escapeXml(comment.author)}"` : ''}><w:p><w:r><w:t>${escapeXml(comment.text)}</w:t></w:r></w:p></w:comment>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${commentXml}</w:comments>`;
}

function createSharedStringPool(workbook: XlsxWorkbook): { values: string[]; indexByValue: Map<string, number> } {
  const values: string[] = [];
  const indexByValue = new Map<string, number>();

  for (const sheet of workbook.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (shouldUseSharedString(cell)) {
          if (!indexByValue.has(cell.value)) {
            indexByValue.set(cell.value, values.length);
            values.push(cell.value);
          }
        }
      }
    }
  }

  return { values, indexByValue };
}

function buildSharedStringsXml(values: string[]): string {
  const items = values.map((value) => `<si><t>${escapeXml(value)}</t></si>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${values.length}" uniqueCount="${values.length}">${items}</sst>`;
}

function buildWorksheetXml(sheet: WorkbookSheet, sharedStringIndices: Map<string, number>, useSharedStrings: boolean): string {
  const rows = sheet.rows.map((row) => `<row r="${row.index}">${row.cells.map((cell) => buildCellXml(cell, sharedStringIndices, useSharedStrings)).join('')}</row>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`;
}

function buildCellXml(cell: WorksheetCell, sharedStringIndices: Map<string, number>, useSharedStrings: boolean): string {
  const styleAttribute = cell.styleIndex !== undefined ? ` s="${cell.styleIndex}"` : '';

  if (cell.formula) {
    return `<c r="${escapeXml(cell.reference)}"${styleAttribute}><f>${escapeXml(cell.formula)}</f><v>${escapeXml(cell.value)}</v></c>`;
  }

  if (shouldUseSharedString(cell)) {
    if (useSharedStrings) {
      const sharedIndex = sharedStringIndices.get(cell.value) ?? 0;
      return `<c r="${escapeXml(cell.reference)}" t="s"${styleAttribute}><v>${sharedIndex}</v></c>`;
    }

    return `<c r="${escapeXml(cell.reference)}" t="inlineStr"${styleAttribute}><is><t>${escapeXml(cell.value)}</t></is></c>`;
  }

  return `<c r="${escapeXml(cell.reference)}"${cell.type !== 'n' ? ` t="${escapeXml(cell.type)}"` : ''}${styleAttribute}><v>${escapeXml(cell.value)}</v></c>`;
}

function shouldUseSharedString(cell: WorksheetCell): boolean {
  if (cell.type === 's') {
    return true;
  }

  return cell.type !== 'n' && !cell.formula && Number.isNaN(Number(cell.value));
}

function buildSlideXml(slide: PresentationSlide): string {
  const shapes = slide.shapes.map((shape) => buildShapeXml(shape)).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>${shapes}</p:spTree></p:cSld></p:sld>`;
}

function buildShapeXml(shape: SlideShape): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${escapeXml(shape.id || '1')}" name="${escapeXml(shape.name ?? 'Shape')}"/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${escapeXml(shape.text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function buildNotesXml(notesText: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>${escapeXml(notesText)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
