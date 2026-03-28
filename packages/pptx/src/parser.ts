import { getParsedXmlPart, relationshipById, relationshipsFor, findElementsByLocalName, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { PresentationComment, PresentationDocument, PresentationSlide, PresentationTheme, PresentationTiming, PresentationTimingNode, PresentationTransition, SlideShape, SlideShapeTransform } from './model';

export function parsePptx(graph: PackageGraph): PresentationDocument {
  const presentationUri = graph.rootDocumentUri ?? '/ppt/presentation.xml';
  const xml = getParsedXmlPart(graph, presentationUri);
  if (!xml) {
    throw new Error('Presentation part is missing.');
  }

  const presentation = xml.document['p:presentation'];
  const slideSize = xmlChild<Record<string, unknown>>(presentation, 'p:sldSz');
  const slidesRoot = xmlChild<Record<string, unknown>>(presentation, 'p:sldIdLst');
  const themeCache: Record<string, PresentationTheme> = {};

  const slides = xmlChildren<Record<string, unknown>>(slidesRoot, 'p:sldId').flatMap((slideReference) => {
    const relationshipId = xmlAttr(slideReference, 'r:id');
    const relationship = relationshipId ? relationshipById(graph, presentationUri, relationshipId) : undefined;
    if (!relationship?.resolvedTarget) {
      return [];
    }

    return [parseSlide(graph, relationship.resolvedTarget, themeCache)];
  });

  return {
    kind: 'pptx',
    packageGraph: graph,
    slides,
    size: {
      cx: Number(xmlAttr(slideSize, 'cx') ?? '0'),
      cy: Number(xmlAttr(slideSize, 'cy') ?? '0')
    },
    themes: themeCache
  };
}

function parseSlide(graph: PackageGraph, uri: string, themeCache: Record<string, PresentationTheme>): PresentationSlide {
  const xml = getParsedXmlPart(graph, uri);
  if (!xml) {
    throw new Error(`Slide part ${uri} is missing.`);
  }

  const slide = xml.document['p:sld'] as Record<string, unknown>;
  const commonSlideData = xmlChild<Record<string, unknown>>(slide, 'p:cSld');
  const shapeTree = xmlChild<Record<string, unknown>>(commonSlideData, 'p:spTree');
  const slideRelationships = relationshipsFor(graph, uri);
  const layoutRelationship = slideRelationships.find((relationship) => relationship.type.includes('/slideLayout'));
  const layoutInfo = layoutRelationship?.resolvedTarget ? parseLayoutInfo(graph, layoutRelationship.resolvedTarget, themeCache) : undefined;
  const shapes = [
    ...xmlChildren<Record<string, unknown>>(shapeTree, 'p:sp').map((shape) => parseShape(shape)),
    ...xmlChildren<Record<string, unknown>>(shapeTree, 'p:pic').map((picture) => parsePicture(picture, slideRelationships)),
    ...xmlChildren<Record<string, unknown>>(shapeTree, 'p:graphicFrame').flatMap((frame) => parseGraphicFrame(frame, slideRelationships))
  ];
  const notesInfo = parseNotesInfo(graph, uri);
  const title = shapes.find((shape) => shape.text.trim())?.text ?? 'Slide';
  const comments = parseSlideComments(graph, uri);

  return {
    title,
    uri,
    notesUri: notesInfo.uri,
    notesText: notesInfo.text,
    layoutUri: layoutInfo?.layoutUri,
    layoutName: layoutInfo?.layoutName,
    masterUri: layoutInfo?.masterUri,
    masterName: layoutInfo?.masterName,
    themeUri: layoutInfo?.themeUri,
    transition: parseTransition(slide),
    timing: parseTiming(slide),
    shapes,
    comments
  };
}

function parseShape(shape: Record<string, unknown>): SlideShape {
  const nvSpPr = xmlChild<Record<string, unknown>>(shape, 'p:nvSpPr');
  const cNvPr = xmlChild<Record<string, unknown>>(nvSpPr, 'p:cNvPr');
  const nvPr = xmlChild<Record<string, unknown>>(nvSpPr, 'p:nvPr');
  const placeholder = xmlChild<Record<string, unknown>>(nvPr, 'p:ph');
  const textNodes = findElementsByLocalName(shape, 't');
  const transform = parseTransform(xmlChild<Record<string, unknown>>(shape, 'p:spPr'));

  return {
    id: xmlAttr(cNvPr, 'id') ?? '',
    name: xmlAttr(cNvPr, 'name'),
    text: textNodes.map((node) => xmlText(node)).join(''),
    placeholderType: xmlAttr(placeholder, 'type'),
    transform
  };
}

function parsePicture(picture: Record<string, unknown>, relationships: ReturnType<typeof relationshipsFor>): SlideShape {
  const nvPicPr = xmlChild<Record<string, unknown>>(picture, 'p:nvPicPr');
  const cNvPr = xmlChild<Record<string, unknown>>(nvPicPr, 'p:cNvPr');
  const blipFill = xmlChild<Record<string, unknown>>(picture, 'p:blipFill');
  const blip = xmlChild<Record<string, unknown>>(blipFill, 'a:blip');
  const relationshipId = xmlAttr(blip, 'r:embed');
  const target = relationshipId ? relationships.find((relationship) => relationship.id === relationshipId)?.resolvedTarget : undefined;
  const transform = parseTransform(xmlChild<Record<string, unknown>>(picture, 'p:spPr'));

  return {
    id: xmlAttr(cNvPr, 'id') ?? '',
    name: xmlAttr(cNvPr, 'name'),
    text: '',
    transform,
    media: {
      type: 'image',
      targetUri: target ?? undefined
    }
  };
}

function parseGraphicFrame(frame: Record<string, unknown>, relationships: ReturnType<typeof relationshipsFor>): SlideShape[] {
  const nvGraphicFramePr = xmlChild<Record<string, unknown>>(frame, 'p:nvGraphicFramePr');
  const cNvPr = xmlChild<Record<string, unknown>>(nvGraphicFramePr, 'p:cNvPr');
  const transform = parseTransform(xmlChild<Record<string, unknown>>(frame, 'p:xfrm'));
  const oleObj = findElementsByLocalName(frame, 'oleObj')[0] as Record<string, unknown> | undefined;
  if (!oleObj) {
    return [];
  }

  const relationshipId = xmlAttr(oleObj, 'r:id');
  const target = relationshipId ? relationships.find((relationship) => relationship.id === relationshipId)?.resolvedTarget : undefined;

  return [{
    id: xmlAttr(cNvPr, 'id') ?? '',
    name: xmlAttr(cNvPr, 'name'),
    text: '',
    transform,
    media: {
      type: 'embeddedObject',
      targetUri: target ?? undefined,
      progId: xmlAttr(oleObj, 'progId') ?? undefined
    }
  }];
}

function parseTransform(shapeProperties: Record<string, unknown> | undefined): SlideShapeTransform | undefined {
  const xfrm = xmlChild<Record<string, unknown>>(shapeProperties, 'a:xfrm') ?? xmlChild<Record<string, unknown>>(shapeProperties, 'p:xfrm');
  if (!xfrm) {
    return undefined;
  }

  const off = xmlChild<Record<string, unknown>>(xfrm, 'a:off');
  const ext = xmlChild<Record<string, unknown>>(xfrm, 'a:ext');
  return {
    x: (() => { const value = xmlAttr(off, 'x'); return value ? Number(value) : undefined; })(),
    y: (() => { const value = xmlAttr(off, 'y'); return value ? Number(value) : undefined; })(),
    cx: (() => { const value = xmlAttr(ext, 'cx'); return value ? Number(value) : undefined; })(),
    cy: (() => { const value = xmlAttr(ext, 'cy'); return value ? Number(value) : undefined; })()
  };
}

function parseLayoutInfo(graph: PackageGraph, layoutUri: string, themeCache: Record<string, PresentationTheme>): { layoutUri: string; layoutName?: string; masterUri?: string; masterName?: string; themeUri?: string } {
  const layoutXml = getParsedXmlPart(graph, layoutUri);
  const layoutRoot = layoutXml?.document['p:sldLayout'] as Record<string, unknown> | undefined;
  const layoutName = layoutRoot ? xmlAttr(layoutRoot, 'matchingName') ?? xmlAttr(layoutRoot, 'type') : undefined;
  const masterRelationship = relationshipsFor(graph, layoutUri).find((relationship) => relationship.type.includes('/slideMaster'));
  const masterUri = masterRelationship?.resolvedTarget ?? undefined;
  let masterName: string | undefined;
  let themeUri: string | undefined;

  if (masterUri) {
    const masterXml = getParsedXmlPart(graph, masterUri);
    const masterRoot = masterXml?.document['p:sldMaster'] as Record<string, unknown> | undefined;
    masterName = masterRoot ? xmlAttr(masterRoot, 'preserve') ?? 'slide-master' : 'slide-master';
    const themeRelationship = relationshipsFor(graph, masterUri).find((relationship) => relationship.type.includes('/theme'));
    themeUri = themeRelationship?.resolvedTarget ?? undefined;
    if (themeUri && !themeCache[themeUri]) {
      themeCache[themeUri] = parseThemeInfo(graph, themeUri);
    }
  }

  return { layoutUri, layoutName, masterUri, masterName, themeUri };
}

function parseThemeInfo(graph: PackageGraph, themeUri: string): PresentationTheme {
  const xml = getParsedXmlPart(graph, themeUri);
  if (!xml) {
    return { uri: themeUri };
  }

  const theme = xml.document['a:theme'] as Record<string, unknown> | undefined;
  const themeElements = xmlChild<Record<string, unknown>>(theme, 'a:themeElements');
  const colorScheme = xmlChild<Record<string, unknown>>(themeElements, 'a:clrScheme');
  const fontScheme = xmlChild<Record<string, unknown>>(themeElements, 'a:fontScheme');
  const majorFont = xmlChild<Record<string, unknown>>(fontScheme, 'a:majorFont');
  const minorFont = xmlChild<Record<string, unknown>>(fontScheme, 'a:minorFont');
  const majorLatin = xmlChild<Record<string, unknown>>(majorFont, 'a:latin');
  const minorLatin = xmlChild<Record<string, unknown>>(minorFont, 'a:latin');

  return {
    uri: themeUri,
    name: theme ? xmlAttr(theme, 'name') : undefined,
    colorSchemeName: colorScheme ? xmlAttr(colorScheme, 'name') : undefined,
    majorLatinFont: majorLatin ? xmlAttr(majorLatin, 'typeface') : undefined,
    minorLatinFont: minorLatin ? xmlAttr(minorLatin, 'typeface') : undefined
  };
}

function parseTransition(slide: Record<string, unknown>): PresentationTransition | undefined {
  const transitionNode = xmlChild<Record<string, unknown>>(slide, 'p:transition');
  if (!transitionNode) {
    return undefined;
  }

  const transitionType = Object.keys(transitionNode).find((key) => !key.startsWith('@_') && key !== '#text');
  return {
    type: transitionType?.split(':').pop(),
    speed: xmlAttr(transitionNode, 'spd'),
    advanceOnClick: xmlAttr(transitionNode, 'advClick') === '1' ? true : xmlAttr(transitionNode, 'advClick') === '0' ? false : undefined,
    advanceAfterMs: (() => { const value = xmlAttr(transitionNode, 'advTm'); return value ? Number(value) : undefined; })()
  };
}

function parseTiming(slide: Record<string, unknown>): PresentationTiming | undefined {
  const timingNode = xmlChild<Record<string, unknown>>(slide, 'p:timing');
  if (!timingNode) {
    return undefined;
  }

  const nodes: PresentationTimingNode[] = [];
  for (const nodeType of ['p:par', 'p:seq', 'p:anim']) {
    for (const node of findElementsByLocalName(timingNode, nodeType.split(':')[1])) {
      const commonTiming = xmlChild<Record<string, unknown>>(node, 'p:cTn');
      const startConditionList = xmlChild<Record<string, unknown>>(node, 'p:stCondLst');
      const startCondition = xmlChild<Record<string, unknown>>(startConditionList, 'p:cond');
      const startTargetElement = xmlChild<Record<string, unknown>>(startCondition, 'p:tgtEl');
      const startShapeTarget = xmlChild<Record<string, unknown>>(startTargetElement, 'p:spTgt');
      const endConditionList = xmlChild<Record<string, unknown>>(node, 'p:endCondLst');
      const endCondition = xmlChild<Record<string, unknown>>(endConditionList, 'p:cond');
      const endTargetElement = xmlChild<Record<string, unknown>>(endCondition, 'p:tgtEl');
      const endShapeTarget = xmlChild<Record<string, unknown>>(endTargetElement, 'p:spTgt');
      const targetElement = xmlChild<Record<string, unknown>>(node, 'p:tgtEl');
      const shapeTarget = xmlChild<Record<string, unknown>>(targetElement, 'p:spTgt');
      nodes.push({
        nodeType: nodeType.split(':')[1],
        concurrent: xmlAttr(node, 'concurrent') === '1' ? true : xmlAttr(node, 'concurrent') === '0' ? false : undefined,
        nextAction: xmlAttr(node, 'nextAc') ?? undefined,
        previousAction: xmlAttr(node, 'prevAc') ?? undefined,
        presetClass: xmlAttr(commonTiming, 'presetClass') ?? xmlAttr(node, 'presetClass') ?? undefined,
        presetId: xmlAttr(commonTiming, 'presetID') ?? xmlAttr(node, 'presetID') ?? undefined,
        id: xmlAttr(commonTiming, 'id') ?? xmlAttr(node, 'id') ?? undefined,
        duration: xmlAttr(commonTiming, 'dur') ?? xmlAttr(node, 'dur') ?? undefined,
        repeatDuration: xmlAttr(commonTiming, 'repeatDur') ?? xmlAttr(node, 'repeatDur') ?? undefined,
        repeatCount: xmlAttr(commonTiming, 'repeatCount') ?? xmlAttr(node, 'repeatCount') ?? undefined,
        restart: xmlAttr(commonTiming, 'restart') ?? xmlAttr(node, 'restart') ?? undefined,
        fill: xmlAttr(commonTiming, 'fill') ?? xmlAttr(node, 'fill') ?? undefined,
        autoReverse: xmlAttr(commonTiming, 'autoRev') === '1' ? true : xmlAttr(commonTiming, 'autoRev') === '0' ? false : xmlAttr(node, 'autoRev') === '1' ? true : xmlAttr(node, 'autoRev') === '0' ? false : undefined,
        acceleration: xmlAttr(commonTiming, 'accel') ?? xmlAttr(node, 'accel') ?? undefined,
        deceleration: xmlAttr(commonTiming, 'decel') ?? xmlAttr(node, 'decel') ?? undefined,
        triggerEvent: xmlAttr(startCondition, 'evt') ?? undefined,
        triggerDelay: xmlAttr(startCondition, 'delay') ?? undefined,
        triggerShapeId: xmlAttr(startShapeTarget, 'spid') ?? undefined,
        endTriggerEvent: xmlAttr(endCondition, 'evt') ?? undefined,
        endTriggerDelay: xmlAttr(endCondition, 'delay') ?? undefined,
        endTriggerShapeId: xmlAttr(endShapeTarget, 'spid') ?? undefined,
        targetShapeId: xmlAttr(shapeTarget, 'spid') ?? undefined
      });
    }
  }

  return {
    nodeCount: nodes.length,
    nodes
  };
}

function parseNotesInfo(graph: PackageGraph, slideUri: string): { uri?: string; text: string } {
  const notesRelationship = relationshipsFor(graph, slideUri).find((relationship) => relationship.type.includes('/notesSlide'));
  if (!notesRelationship?.resolvedTarget) {
    return { text: '' };
  }

  const xml = getParsedXmlPart(graph, notesRelationship.resolvedTarget);
  if (!xml) {
    return { uri: notesRelationship.resolvedTarget, text: '' };
  }

  return {
    uri: notesRelationship.resolvedTarget,
    text: findElementsByLocalName(xml.document, 't').map((node) => xmlText(node)).join('')
  };
}

function parseSlideComments(graph: PackageGraph, slideUri: string): PresentationComment[] {
  const commentsRelationship = relationshipsFor(graph, slideUri).find((relationship) => relationship.type.includes('/comments'));
  if (!commentsRelationship?.resolvedTarget) {
    return [];
  }

  const xml = getParsedXmlPart(graph, commentsRelationship.resolvedTarget);
  if (!xml) {
    return [];
  }

  const root = xml.document['p:cmLst'];
  return xmlChildren<Record<string, unknown>>(root, 'p:cm').map((commentNode, index) => ({
    author: xmlAttr(commentNode, 'authorId') ?? undefined,
    text: xmlText(xmlChild(commentNode, 'p:text')),
    index
  }));
}
