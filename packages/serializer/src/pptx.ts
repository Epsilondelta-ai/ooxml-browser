import { clonePackageGraph, relationshipsFor, replaceInnerTextByAttribute, serializePackageGraph, updatePackagePartText } from '@ooxml/core';
import type { PresentationComment, PresentationDocument, PresentationSlide, PresentationTimingNode, SlideShape } from '@ooxml/pptx';

export function serializePptx(presentation: PresentationDocument): Uint8Array {
  const graph = clonePackageGraph(presentation.packageGraph);

  for (const slide of presentation.slides) {
    updatePackagePartText(
      graph,
      slide.uri,
      buildSlideXml(slide, relationshipsFor(graph, slide.uri)),
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

    const commentsRelationship = relationshipsFor(graph, slide.uri).find((relationship) => relationship.type.includes('/comments'));
    if (commentsRelationship?.resolvedTarget) {
      updatePackagePartText(
        graph,
        commentsRelationship.resolvedTarget,
        buildCommentsXml(slide.comments, graph.parts[commentsRelationship.resolvedTarget]?.text),
        'application/vnd.openxmlformats-officedocument.presentationml.comment+xml'
      );
    }
  }

  return serializePackageGraph(graph);
}

function buildSlideXml(slide: PresentationSlide, slideRelationships: ReturnType<typeof relationshipsFor>): string {
  const shapes = slide.shapes.map((shape) => buildShapeXml(shape, slideRelationships)).join('');
  const transitionXml = slide.transition ? buildTransitionXml(slide.transition.type, slide.transition.speed) : '';
  const timingXml = slide.timing ? buildTimingXml(slide.timing.nodes) : '';
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${shapes}</p:spTree></p:cSld>${transitionXml}${timingXml}</p:sld>`;
}

function buildShapeXml(shape: SlideShape, slideRelationships: ReturnType<typeof relationshipsFor>): string {
  const transformXml = buildTransformXml(shape.transform);

  if (shape.media?.type === 'image') {
    const relationshipId = slideRelationships.find((relationship) => relationship.resolvedTarget === shape.media?.targetUri)?.id ?? 'rIdImage';
    return `<p:pic><p:nvPicPr><p:cNvPr id="${escapeXml(shape.id || '1')}" name="${escapeXml(shape.name ?? 'Image')}"/></p:nvPicPr><p:spPr>${transformXml}</p:spPr><p:blipFill><a:blip r:embed="${escapeXml(relationshipId)}"/></p:blipFill></p:pic>`;
  }

  return `<p:sp><p:nvSpPr><p:cNvPr id="${escapeXml(shape.id || '1')}" name="${escapeXml(shape.name ?? 'Shape')}"/><p:nvPr>${shape.placeholderType ? `<p:ph type="${escapeXml(shape.placeholderType)}"/>` : ''}</p:nvPr></p:nvSpPr><p:spPr>${transformXml}</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${escapeXml(shape.text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function buildTransformXml(transform: SlideShape['transform']): string {
  if (!transform || [transform.x, transform.y, transform.cx, transform.cy].every((value) => value === undefined)) {
    return '';
  }

  return `<a:xfrm><a:off x="${transform.x ?? 0}" y="${transform.y ?? 0}"/><a:ext cx="${transform.cx ?? 0}" cy="${transform.cy ?? 0}"/></a:xfrm>`;
}

function buildTransitionXml(type: string | undefined, speed: string | undefined): string {
  if (!type) {
    return '';
  }

  return `<p:transition${speed ? ` spd="${escapeXml(speed)}"` : ''}><p:${escapeXml(type)}/></p:transition>`;
}

function buildTimingXml(nodes: PresentationTimingNode[]): string {
  if (nodes.length === 0) {
    return '';
  }

  const body = nodes.map((node) => `<p:${node.nodeType}><p:cTn${node.presetClass ? ` presetClass="${escapeXml(node.presetClass)}"` : ''}${node.presetId ? ` presetID="${escapeXml(node.presetId)}"` : ''}/></p:${node.nodeType}>`).join('');
  return `<p:timing><p:tnLst>${body}</p:tnLst></p:timing>`;
}

function buildCommentsXml(comments: PresentationComment[], existingSource?: string): string {
  if (existingSource) {
    let next = existingSource;
    comments.forEach((comment, index) => {
      next = replaceInnerTextByAttribute(next, { containerTag: 'p:cm', occurrence: index, textTag: 'p:text', newText: comment.text });
    });
    return next;
  }

  const body = comments.map((comment) => `<p:cm${comment.author ? ` authorId="${escapeXml(comment.author)}"` : ''}><p:text>${escapeXml(comment.text)}</p:text></p:cm>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${body}</p:cmLst>`;
}

function buildNotesXml(notesText: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>${escapeXml(notesText)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
