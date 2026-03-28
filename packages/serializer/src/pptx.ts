import { clonePackageGraph, serializePackageGraph, updatePackagePartText } from '@ooxml/core';
import type { PresentationDocument, PresentationSlide, SlideShape } from '@ooxml/pptx';

export function serializePptx(presentation: PresentationDocument): Uint8Array {
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
