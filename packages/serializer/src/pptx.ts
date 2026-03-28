import { applyXmlPatchPlan, clonePackageGraph, relationshipsFor, serializePackageGraph, updatePackagePartText, upsertRelationship } from '@ooxml/core';
import { parsePptx, type PresentationComment, type PresentationDocument, type PresentationSlide, type PresentationTimingNode, type SlideShape } from '@ooxml/pptx';

export function serializePptx(presentation: PresentationDocument): Uint8Array {
  const graph = clonePackageGraph(presentation.packageGraph);
  const originalPresentation = parsePptx(presentation.packageGraph);
  const originalSlidesByUri = new Map(originalPresentation.slides.map((slide) => [slide.uri, slide]));

  const presentationUri = presentation.packageGraph.rootDocumentUri ?? '/ppt/presentation.xml';
  if (graph.parts[presentationUri]) {
    const existingPresentationSource = graph.parts[presentationUri]?.text;
    updatePackagePartText(
      graph,
      presentationUri,
      existingPresentationSource ? patchPresentationXml(existingPresentationSource, presentation.size) : buildPresentationXml(presentation),
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
    );
  }

  for (const slide of presentation.slides) {
    const originalSlide = originalSlidesByUri.get(slide.uri);
    syncSlideLayoutRelationship(graph, slide.uri, originalSlide, slide);
    syncSlideMasterRelationship(graph, originalSlide, slide);
    syncSlideThemeRelationship(graph, originalSlide, slide);
    syncSlideImageRelationships(graph, slide.uri, originalSlide, slide);
    const existingSlideSource = graph.parts[slide.uri]?.text;
    const nextSlideSource =
      originalSlide && existingSlideSource
        ? patchSlideXml(existingSlideSource, originalSlide, slide) ?? buildSlideXml(slide, relationshipsFor(graph, slide.uri))
        : buildSlideXml(slide, relationshipsFor(graph, slide.uri));

    updatePackagePartText(
      graph,
      slide.uri,
      nextSlideSource,
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    );

    const notesUri = slide.notesUri;
    if (notesUri && graph.parts[notesUri]) {
      const existingNotesSource = graph.parts[notesUri]?.text;
      updatePackagePartText(
        graph,
        notesUri,
        existingNotesSource ? patchNotesXml(existingNotesSource, slide.notesText) : buildNotesXml(slide.notesText),
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

function syncSlideLayoutRelationship(graph: PresentationDocument['packageGraph'], slideUri: string, originalSlide: PresentationSlide | undefined, slide: PresentationSlide): void {
  if (!slide.layoutUri || slide.layoutUri === originalSlide?.layoutUri) {
    return;
  }

  const slideRelationships = relationshipsFor(graph, slideUri);
  const layoutRelationship = slideRelationships.find((relationship) => relationship.type.includes('/slideLayout'));
  if (!layoutRelationship) {
    return;
  }

  upsertRelationship(graph, slideUri, {
    id: layoutRelationship.id,
    type: layoutRelationship.type,
    target: relativeRelationshipTarget(slideUri, slide.layoutUri),
    targetMode: 'Internal'
  });
}

function syncSlideThemeRelationship(graph: PresentationDocument['packageGraph'], originalSlide: PresentationSlide | undefined, slide: PresentationSlide): void {
  const masterUri = slide.masterUri ?? originalSlide?.masterUri;
  if (!masterUri || !slide.themeUri || slide.themeUri === originalSlide?.themeUri) {
    return;
  }

  const masterRelationships = relationshipsFor(graph, masterUri);
  const themeRelationship = masterRelationships.find((relationship) => relationship.type.includes('/theme'));
  if (!themeRelationship) {
    return;
  }

  upsertRelationship(graph, masterUri, {
    id: themeRelationship.id,
    type: themeRelationship.type,
    target: relativeRelationshipTarget(masterUri, slide.themeUri),
    targetMode: 'Internal'
  });
}

function syncSlideMasterRelationship(graph: PresentationDocument['packageGraph'], originalSlide: PresentationSlide | undefined, slide: PresentationSlide): void {
  if (!slide.layoutUri || !slide.masterUri || slide.masterUri === originalSlide?.masterUri) {
    return;
  }

  const layoutRelationships = relationshipsFor(graph, slide.layoutUri);
  const masterRelationship = layoutRelationships.find((relationship) => relationship.type.includes('/slideMaster'));
  if (!masterRelationship) {
    return;
  }

  upsertRelationship(graph, slide.layoutUri, {
    id: masterRelationship.id,
    type: masterRelationship.type,
    target: relativeRelationshipTarget(slide.layoutUri, slide.masterUri),
    targetMode: 'Internal'
  });
}

function syncSlideImageRelationships(graph: PresentationDocument['packageGraph'], slideUri: string, originalSlide: PresentationSlide | undefined, slide: PresentationSlide): void {
  const slideRelationships = relationshipsFor(graph, slideUri);
  const imageRelationships = slideRelationships.filter((relationship) => relationship.type.includes('/image'));
  let imageRelationshipIndex = 0;

  for (const [shapeIndex, shape] of slide.shapes.entries()) {
    if (shape.media?.type !== 'image' || !shape.media.targetUri) {
      continue;
    }

    const originalShape = originalSlide?.shapes[shapeIndex];
    const relationship =
      (originalShape?.media?.targetUri
        ? imageRelationships.find((entry) => entry.resolvedTarget === originalShape.media?.targetUri)
        : undefined)
      ?? imageRelationships[imageRelationshipIndex];

    imageRelationshipIndex += 1;
    if (!relationship || relationship.resolvedTarget === shape.media.targetUri) {
      continue;
    }

    upsertRelationship(graph, slideUri, {
      id: relationship.id,
      type: relationship.type,
      target: relativeRelationshipTarget(slideUri, shape.media.targetUri),
      targetMode: 'Internal'
    });
  }
}

function buildPresentationXml(presentation: PresentationDocument): string {
  const slideEntries = presentation.slides.map((_slide, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="${presentation.size.cx}" cy="${presentation.size.cy}"/><p:sldIdLst>${slideEntries}</p:sldIdLst></p:presentation>`;
}

function patchPresentationXml(existingSource: string, size: { cx: number; cy: number }): string {
  return applyXmlPatchPlan(existingSource, [
    { op: 'replaceAttribute', tagName: 'p:sldSz', targetAttr: 'cx', newValue: String(size.cx) },
    { op: 'replaceAttribute', tagName: 'p:sldSz', targetAttr: 'cy', newValue: String(size.cy) }
  ]);
}

function patchSlideXml(existingSource: string, originalSlide: PresentationSlide, slide: PresentationSlide): string | undefined {
  if (!canPatchSlideTextOnly(originalSlide, slide)) {
    return undefined;
  }

  const currentTextShapes = slide.shapes.filter((shape) => shape.media?.type !== 'image');
  const originalTextShapes = originalSlide.shapes.filter((shape) => shape.media?.type !== 'image');
  if (currentTextShapes.length !== originalTextShapes.length) {
    return undefined;
  }

  const operations = currentTextShapes.flatMap((shape, index) => {
    if (shape.text === originalTextShapes[index]?.text) {
      return [];
    }

    return [{
      op: 'replaceText' as const,
      containerTag: 'p:sp',
      occurrence: index,
      textTag: 'a:t',
      newText: shape.text
    }];
  });

  return operations.length > 0 ? applyXmlPatchPlan(existingSource, operations) : existingSource;
}

function canPatchSlideTextOnly(originalSlide: PresentationSlide, slide: PresentationSlide): boolean {
  if (slide.shapes.length !== originalSlide.shapes.length) {
    return false;
  }

  if (JSON.stringify(slide.transition ?? null) !== JSON.stringify(originalSlide.transition ?? null)) {
    return false;
  }

  if (JSON.stringify(slide.timing ?? null) !== JSON.stringify(originalSlide.timing ?? null)) {
    return false;
  }

  return slide.shapes.every((shape, index) => {
    const originalShape = originalSlide.shapes[index];
    if (!originalShape) {
      return false;
    }

    return shape.id === originalShape.id
      && shape.name === originalShape.name
      && shape.placeholderType === originalShape.placeholderType
      && JSON.stringify(shape.media ?? null) === JSON.stringify(originalShape.media ?? null)
      && JSON.stringify(shape.transform ?? null) === JSON.stringify(originalShape.transform ?? null);
  });
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

  const body = nodes.map((node) => {
    const conditionXml = node.triggerEvent || node.triggerDelay
      ? `<p:stCondLst><p:cond${node.triggerEvent ? ` evt="${escapeXml(node.triggerEvent)}"` : ''}${node.triggerDelay ? ` delay="${escapeXml(node.triggerDelay)}"` : ''}/></p:stCondLst>`
      : '';
    const endConditionXml = node.endTriggerEvent || node.endTriggerDelay
      ? `<p:endCondLst><p:cond${node.endTriggerEvent ? ` evt="${escapeXml(node.endTriggerEvent)}"` : ''}${node.endTriggerDelay ? ` delay="${escapeXml(node.endTriggerDelay)}"` : ''}/></p:endCondLst>`
      : '';
    const targetXml = node.targetShapeId
      ? `<p:tgtEl><p:spTgt spid="${escapeXml(node.targetShapeId)}"/></p:tgtEl>`
      : '';
    return `<p:${node.nodeType}><p:cTn${node.id ? ` id="${escapeXml(node.id)}"` : ''}${node.presetClass ? ` presetClass="${escapeXml(node.presetClass)}"` : ''}${node.presetId ? ` presetID="${escapeXml(node.presetId)}"` : ''}${node.duration ? ` dur="${escapeXml(node.duration)}"` : ''}${node.repeatDuration ? ` repeatDur="${escapeXml(node.repeatDuration)}"` : ''}${node.repeatCount ? ` repeatCount="${escapeXml(node.repeatCount)}"` : ''}${node.restart ? ` restart="${escapeXml(node.restart)}"` : ''}${node.fill ? ` fill="${escapeXml(node.fill)}"` : ''}${node.autoReverse !== undefined ? ` autoRev="${node.autoReverse ? '1' : '0'}"` : ''}${node.acceleration ? ` accel="${escapeXml(node.acceleration)}"` : ''}${node.deceleration ? ` decel="${escapeXml(node.deceleration)}"` : ''}/>${conditionXml}${endConditionXml}${targetXml}</p:${node.nodeType}>`;
  }).join('');
  return `<p:timing><p:tnLst>${body}</p:tnLst></p:timing>`;
}

function buildCommentsXml(comments: PresentationComment[], existingSource?: string): string {
  const existingCount = existingSource ? (existingSource.match(/<p:cm\b/g) ?? []).length : 0;
  if (existingSource && existingCount === comments.length && comments.length > 0) {
    return applyXmlPatchPlan(existingSource, comments.flatMap((comment, index) => [
      {
        op: 'replaceAttribute' as const,
        tagName: 'p:cm',
        targetAttr: 'authorId',
        newValue: comment.author ?? '',
        occurrence: index
      },
      {
        op: 'replaceText' as const,
        containerTag: 'p:cm',
        occurrence: index,
        textTag: 'p:text',
        newText: comment.text
      }
    ]));
  }

  const body = comments.map((comment) => `<p:cm${comment.author ? ` authorId="${escapeXml(comment.author)}"` : ''}><p:text>${escapeXml(comment.text)}</p:text></p:cm>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${body}</p:cmLst>`;
}

function patchNotesXml(existingSource: string, notesText: string): string {
  return applyXmlPatchPlan(existingSource, [{
    op: 'replaceText',
    containerTag: 'p:sp',
    occurrence: 0,
    textTag: 'a:t',
    newText: notesText
  }]);
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

function relativeRelationshipTarget(sourceUri: string, targetUri: string): string {
  const sourceSegments = sourceUri.replace(/^\//, '').split('/');
  sourceSegments.pop();
  const targetSegments = targetUri.replace(/^\//, '').split('/');

  while (sourceSegments.length > 0 && targetSegments.length > 0 && sourceSegments[0] === targetSegments[0]) {
    sourceSegments.shift();
    targetSegments.shift();
  }

  return `${sourceSegments.map(() => '..').join('/')}${sourceSegments.length ? '/' : ''}${targetSegments.join('/')}`;
}
