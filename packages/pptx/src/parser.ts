import { getParsedXmlPart, relationshipById, relationshipsFor, findElementsByLocalName, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { PresentationComment, PresentationDocument, PresentationFill, PresentationLine, PresentationPathCommand, PresentationSlide, PresentationTextStyle, PresentationTheme, PresentationTiming, PresentationTimingNode, PresentationTransition, SlideShape, SlideShapeTransform } from './model';

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
  const rawShapeXmlById = buildRawShapeXmlIndex(xml.source);
  const layoutRelationship = slideRelationships.find((relationship) => relationship.type.includes('/slideLayout'));
  const layoutInfo = layoutRelationship?.resolvedTarget ? parseLayoutInfo(graph, layoutRelationship.resolvedTarget, themeCache) : undefined;
  const theme = layoutInfo?.themeUri ? themeCache[layoutInfo.themeUri] : undefined;
  const baseShapes = collectShapes(shapeTree, slideRelationships, theme, undefined, rawShapeXmlById);
  const inheritedShapes = [
    ...(layoutInfo?.masterShapes ?? []),
    ...(layoutInfo?.layoutShapes ?? [])
  ];
  const shapes = [
    ...inheritedShapes.filter((shape) => !shape.placeholderType && !shape.placeholderIndex),
    ...applyInheritedPlaceholderProperties(baseShapes, inheritedShapes)
  ];
  const notesInfo = parseNotesInfo(graph, uri);
  const title = selectSlideTitle(shapes);
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
    background: parseBackground(commonSlideData, slideRelationships, theme) ?? layoutInfo?.layoutBackground ?? layoutInfo?.masterBackground,
    transition: parseTransition(slide),
    timing: parseTiming(slide),
    shapes,
    comments
  };
}

function selectSlideTitle(shapes: SlideShape[]): string {
  const candidates = shapes
    .map((shape) => ({ shape, text: shape.text.trim() }))
    .filter((entry) => entry.text.length > 0);

  if (!candidates.length) {
    return 'Slide';
  }

  const maxBottom = Math.max(
    ...candidates.map(({ shape }) => {
      const y = shape.transform?.y ?? 0;
      const cy = shape.transform?.cy ?? 0;
      return y + cy;
    }),
    1
  );

  const scored = candidates.map(({ shape, text }) => {
    const y = shape.transform?.y ?? 0;
    const cx = shape.transform?.cx ?? Math.max(text.length * 10_000, 1);
    const cy = shape.transform?.cy ?? Math.max(text.length * 2_000, 1);
    const bottomRatio = (y + cy) / maxBottom;
    let score = 0;

    if (shape.placeholderType === 'title' || shape.placeholderType === 'ctrTitle') {
      score += 10_000;
    }

    if (/^https?:\/\//i.test(text)) {
      score -= 8_000;
    }

    if (bottomRatio > 0.88) {
      score -= 4_000;
    }

    if (text.length < 4) {
      score -= 1_000;
    }

    score += Math.min((cx * cy) / 50_000_000, 6_000);
    score -= y / 2_000;

    return { text, score };
  });

  return scored.sort((left, right) => right.score - left.score)[0]?.text ?? 'Slide';
}

interface TransformContext {
  a: number;
  b: number;
  c: number;
  d: number;
  tx: number;
  ty: number;
  rotationDeg: number;
  flipH: boolean;
  flipV: boolean;
}

interface PathGeometry {
  commands?: PresentationPathCommand[];
  viewport?: {
    width: number;
    height: number;
  };
}

const identityTransformContext: TransformContext = {
  a: 1,
  b: 0,
  c: 0,
  d: 1,
  tx: 0,
  ty: 0,
  rotationDeg: 0,
  flipH: false,
  flipV: false
};

function collectShapes(
  container: Record<string, unknown> | undefined,
  relationships: ReturnType<typeof relationshipsFor>,
  theme?: PresentationTheme,
  context: TransformContext = identityTransformContext,
  rawShapeXmlById?: Record<string, string>
): SlideShape[] {
  if (!container) {
    return [];
  }

  return [
    ...xmlChildren<Record<string, unknown>>(container, 'p:sp').map((shape) => parseShape(shape, relationships, theme, context, rawShapeXmlById)),
    ...xmlChildren<Record<string, unknown>>(container, 'p:pic').map((picture) => parsePicture(picture, relationships, theme, context, rawShapeXmlById)),
    ...xmlChildren<Record<string, unknown>>(container, 'p:graphicFrame').flatMap((frame) => parseGraphicFrame(frame, relationships, theme, context)),
    ...xmlChildren<Record<string, unknown>>(container, 'p:grpSp').flatMap((group) => collectShapes(group, relationships, theme, composeGroupContext(group, context), rawShapeXmlById))
  ];
}

function parseShape(
  shape: Record<string, unknown>,
  relationships: ReturnType<typeof relationshipsFor>,
  theme?: PresentationTheme,
  context?: TransformContext,
  rawShapeXmlById?: Record<string, string>
): SlideShape {
  const nvSpPr = xmlChild<Record<string, unknown>>(shape, 'p:nvSpPr');
  const cNvPr = xmlChild<Record<string, unknown>>(nvSpPr, 'p:cNvPr');
  const nvPr = xmlChild<Record<string, unknown>>(nvSpPr, 'p:nvPr');
  const placeholder = xmlChild<Record<string, unknown>>(nvPr, 'p:ph');
  const shapeProperties = xmlChild<Record<string, unknown>>(shape, 'p:spPr');
  const style = xmlChild<Record<string, unknown>>(shape, 'p:style');
  const shapeType = parseShapeType(shapeProperties);
  const transform = applyTransformContext(parseTransform(shapeProperties), context);
  const pathGeometry = parsePathGeometry(shapeProperties, rawShapeXmlById?.[xmlAttr(cNvPr, 'id') ?? '']);

  return {
    id: xmlAttr(cNvPr, 'id') ?? '',
    name: xmlAttr(cNvPr, 'name'),
    text: extractShapeText(shape),
    placeholderType: xmlAttr(placeholder, 'type'),
    placeholderIndex: xmlAttr(placeholder, 'idx') ?? undefined,
    shapeType,
    transform,
    fill: parseFill(shapeProperties, relationships, theme, style, shapeType),
    line: parseLine(shapeProperties, theme, style),
    textStyle: parseTextStyle(shape, theme),
    pathCommands: pathGeometry.commands,
    pathViewport: pathGeometry.viewport
  };
}

function extractShapeText(shape: Record<string, unknown>): string {
  const textBody = xmlChild<Record<string, unknown>>(shape, 'p:txBody');
  if (!textBody) {
    return '';
  }

  return xmlChildren<Record<string, unknown>>(textBody, 'a:p')
    .map((paragraph) => findElementsByLocalName(paragraph, 't').map((node) => xmlText(node)).join(''))
    .filter((text) => text.length > 0)
    .join('\n');
}

function parsePicture(
  picture: Record<string, unknown>,
  relationships: ReturnType<typeof relationshipsFor>,
  theme?: PresentationTheme,
  context?: TransformContext,
  rawShapeXmlById?: Record<string, string>
): SlideShape {
  const nvPicPr = xmlChild<Record<string, unknown>>(picture, 'p:nvPicPr');
  const cNvPr = xmlChild<Record<string, unknown>>(nvPicPr, 'p:cNvPr');
  const blipFill = xmlChild<Record<string, unknown>>(picture, 'p:blipFill');
  const blip = xmlChild<Record<string, unknown>>(blipFill, 'a:blip');
  const relationshipId = xmlAttr(blip, 'r:embed');
  const target = relationshipId ? relationships.find((relationship) => relationship.id === relationshipId)?.resolvedTarget : undefined;
  const shapeProperties = xmlChild<Record<string, unknown>>(picture, 'p:spPr');
  const style = xmlChild<Record<string, unknown>>(picture, 'p:style');
  const transform = applyTransformContext(parseTransform(shapeProperties), context);
  const pathGeometry = parsePathGeometry(shapeProperties, rawShapeXmlById?.[xmlAttr(cNvPr, 'id') ?? '']);

  return {
    id: xmlAttr(cNvPr, 'id') ?? '',
    name: xmlAttr(cNvPr, 'name'),
    text: '',
    shapeType: 'picture',
    transform,
    fill: parseFill(shapeProperties, relationships, theme, style, 'picture'),
    line: parseLine(shapeProperties, theme, style),
    pathCommands: pathGeometry.commands,
    pathViewport: pathGeometry.viewport,
    media: {
      type: 'image',
      targetUri: target ?? undefined
    }
  };
}

function parseGraphicFrame(frame: Record<string, unknown>, relationships: ReturnType<typeof relationshipsFor>, _theme?: PresentationTheme, context?: TransformContext): SlideShape[] {
  const nvGraphicFramePr = xmlChild<Record<string, unknown>>(frame, 'p:nvGraphicFramePr');
  const cNvPr = xmlChild<Record<string, unknown>>(nvGraphicFramePr, 'p:cNvPr');
  const transform = applyTransformContext(parseTransform(xmlChild<Record<string, unknown>>(frame, 'p:xfrm')), context);
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
    shapeType: 'graphicFrame',
    transform,
    fill: undefined,
    line: undefined,
    pathCommands: undefined,
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
    cy: (() => { const value = xmlAttr(ext, 'cy'); return value ? Number(value) : undefined; })(),
    rotationDeg: (() => { const value = xmlAttr(xfrm, 'rot'); return value ? Number(value) / 60_000 : undefined; })(),
    flipH: xmlAttr(xfrm, 'flipH') === '1' || undefined,
    flipV: xmlAttr(xfrm, 'flipV') === '1' || undefined
  };
}

function composeGroupContext(group: Record<string, unknown>, parent: TransformContext): TransformContext {
  const groupProperties = xmlChild<Record<string, unknown>>(group, 'p:grpSpPr');
  const xfrm = xmlChild<Record<string, unknown>>(groupProperties, 'a:xfrm');
  const off = xmlChild<Record<string, unknown>>(xfrm, 'a:off');
  const ext = xmlChild<Record<string, unknown>>(xfrm, 'a:ext');
  const chOff = xmlChild<Record<string, unknown>>(xfrm, 'a:chOff');
  const chExt = xmlChild<Record<string, unknown>>(xfrm, 'a:chExt');

  const offsetX = Number(xmlAttr(off, 'x') ?? 0);
  const offsetY = Number(xmlAttr(off, 'y') ?? 0);
  const childOffsetX = Number(xmlAttr(chOff, 'x') ?? 0);
  const childOffsetY = Number(xmlAttr(chOff, 'y') ?? 0);
  const extX = Number(xmlAttr(ext, 'cx') ?? 0);
  const extY = Number(xmlAttr(ext, 'cy') ?? 0);
  const childExtX = Number(xmlAttr(chExt, 'cx') ?? 0);
  const childExtY = Number(xmlAttr(chExt, 'cy') ?? 0);
  const scaleX = childExtX ? extX / childExtX : 1;
  const scaleY = childExtY ? extY / childExtY : 1;
  const rotationDeg = Number(xmlAttr(xfrm, 'rot') ?? 0) / 60_000;
  const flipH = xmlAttr(xfrm, 'flipH') === '1';
  const flipV = xmlAttr(xfrm, 'flipV') === '1';

  const localMatrix = multiplyMatrices(
    translationMatrix(offsetX, offsetY),
    multiplyMatrices(
      aroundPointMatrix(extX / 2, extY / 2, multiplyMatrices(rotationMatrix(rotationDeg), scaleMatrix(flipH ? -1 : 1, flipV ? -1 : 1))),
      multiplyMatrices(scaleMatrix(scaleX, scaleY), translationMatrix(-childOffsetX, -childOffsetY))
    )
  );
  const matrix = multiplyMatrices(parent, localMatrix);

  return {
    ...matrix,
    rotationDeg: parent.rotationDeg + rotationDeg,
    flipH: parent.flipH !== flipH,
    flipV: parent.flipV !== flipV
  };
}

function applyTransformContext(
  transform: SlideShapeTransform | undefined,
  context: TransformContext = identityTransformContext
): SlideShapeTransform | undefined {
  if (!transform) {
    return undefined;
  }

  const topLeft = transform.x !== undefined && transform.y !== undefined
    ? applyMatrixToPoint(context, transform.x, transform.y)
    : undefined;
  const widthVector = transform.cx !== undefined ? applyMatrixToVector(context, transform.cx, 0) : undefined;
  const heightVector = transform.cy !== undefined ? applyMatrixToVector(context, 0, transform.cy) : undefined;
  const width = widthVector ? Math.hypot(widthVector.x, widthVector.y) : undefined;
  const height = heightVector ? Math.hypot(heightVector.x, heightVector.y) : undefined;
  const center = (
    transform.x !== undefined
    && transform.y !== undefined
    && transform.cx !== undefined
    && transform.cy !== undefined
  )
    ? applyMatrixToPoint(context, transform.x + transform.cx / 2, transform.y + transform.cy / 2)
    : undefined;

  return {
    x: center && width !== undefined ? center.x - width / 2 : topLeft?.x,
    y: center && height !== undefined ? center.y - height / 2 : topLeft?.y,
    cx: width,
    cy: height,
    rotationDeg: (transform.rotationDeg ?? 0) + context.rotationDeg || undefined,
    flipH: context.flipH !== Boolean(transform.flipH) || undefined,
    flipV: context.flipV !== Boolean(transform.flipV) || undefined
  };
}

function multiplyMatrices(left: TransformContext, right: Pick<TransformContext, 'a' | 'b' | 'c' | 'd' | 'tx' | 'ty'>): TransformContext {
  return {
    ...identityTransformContext,
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    tx: left.a * right.tx + left.c * right.ty + left.tx,
    ty: left.b * right.tx + left.d * right.ty + left.ty
  };
}

function translationMatrix(tx: number, ty: number): TransformContext {
  return {
    ...identityTransformContext,
    tx,
    ty
  };
}

function scaleMatrix(scaleX: number, scaleY: number): TransformContext {
  return {
    ...identityTransformContext,
    a: scaleX,
    d: scaleY
  };
}

function rotationMatrix(rotationDeg: number): TransformContext {
  const radians = rotationDeg * (Math.PI / 180);
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  return {
    ...identityTransformContext,
    a: cos,
    b: sin,
    c: -sin,
    d: cos
  };
}

function aroundPointMatrix(cx: number, cy: number, matrix: TransformContext): TransformContext {
  return multiplyMatrices(
    translationMatrix(cx, cy),
    multiplyMatrices(matrix, translationMatrix(-cx, -cy))
  );
}

function applyMatrixToPoint(context: TransformContext, x: number, y: number): { x: number; y: number } {
  return {
    x: context.a * x + context.c * y + context.tx,
    y: context.b * x + context.d * y + context.ty
  };
}

function applyMatrixToVector(context: TransformContext, x: number, y: number): { x: number; y: number } {
  return {
    x: context.a * x + context.c * y,
    y: context.b * x + context.d * y
  };
}

function parseShapeType(shapeProperties: Record<string, unknown> | undefined): string | undefined {
  if (xmlChild<Record<string, unknown>>(shapeProperties, 'a:custGeom')) {
    return 'custom';
  }

  const preset = xmlChild<Record<string, unknown>>(shapeProperties, 'a:prstGeom');
  return xmlAttr(preset, 'prst') ?? undefined;
}

function parsePathGeometry(shapeProperties: Record<string, unknown> | undefined, rawShapeXml?: string): PathGeometry {
  const orderedCommands = rawShapeXml && (rawShapeXml.includes('<a:gradFill') || rawShapeXml.includes('txBox="1"'))
    ? parseOrderedPathGeometry(rawShapeXml)
    : undefined;
  if (orderedCommands?.commands?.length) {
    return orderedCommands;
  }
  const customGeometry = xmlChild<Record<string, unknown>>(shapeProperties, 'a:custGeom');
  const pathList = xmlChild<Record<string, unknown>>(customGeometry, 'a:pathLst');
  const pathNodes = xmlChildren<Record<string, unknown>>(pathList, 'a:path');
  if (!pathNodes.length) {
    return {};
  }

  const commands: PresentationPathCommand[] = [];
  let viewportWidth = 0;
  let viewportHeight = 0;
  for (const pathNode of pathNodes) {
    viewportWidth = Math.max(viewportWidth, Number(xmlAttr(pathNode, 'w') ?? 0));
    viewportHeight = Math.max(viewportHeight, Number(xmlAttr(pathNode, 'h') ?? 0));
    for (const [key, value] of Object.entries(pathNode)) {
      if (key.startsWith('@_')) {
        continue;
      }

      const entries = Array.isArray(value) ? value : [value];
      for (const entry of entries) {
        if (key === 'a:moveTo' || key === 'a:lnTo') {
          const point = xmlChild<Record<string, unknown>>(entry, 'a:pt');
          commands.push({
            type: key === 'a:moveTo' ? 'moveTo' : 'lineTo',
            x: (() => { const raw = xmlAttr(point, 'x'); return raw ? Number(raw) : undefined; })(),
            y: (() => { const raw = xmlAttr(point, 'y'); return raw ? Number(raw) : undefined; })()
          });
        } else if (key === 'a:cubicBezTo') {
          const points = xmlChildren<Record<string, unknown>>(entry, 'a:pt');
          commands.push({
            type: 'cubicTo',
            x1: (() => { const raw = xmlAttr(points[0], 'x'); return raw ? Number(raw) : undefined; })(),
            y1: (() => { const raw = xmlAttr(points[0], 'y'); return raw ? Number(raw) : undefined; })(),
            x2: (() => { const raw = xmlAttr(points[1], 'x'); return raw ? Number(raw) : undefined; })(),
            y2: (() => { const raw = xmlAttr(points[1], 'y'); return raw ? Number(raw) : undefined; })(),
            x: (() => { const raw = xmlAttr(points[2], 'x'); return raw ? Number(raw) : undefined; })(),
            y: (() => { const raw = xmlAttr(points[2], 'y'); return raw ? Number(raw) : undefined; })()
          });
        } else if (key === 'a:close') {
          commands.push({ type: 'close' });
        }
      }
    }
  }

  return {
    commands: commands.length ? commands : undefined,
    viewport: viewportWidth > 0 && viewportHeight > 0
      ? { width: viewportWidth, height: viewportHeight }
      : undefined
  };
}

function buildRawShapeXmlIndex(source: string): Record<string, string> {
  const entries: Record<string, string> = {};
  for (const tag of ['sp', 'pic']) {
    const pattern = new RegExp(`<p:${tag}\\b[\\s\\S]*?<p:cNvPr[^>]*\\bid="([^"]+)"[^>]*>[\\s\\S]*?<\\/p:${tag}>`, 'g');
    for (const match of source.matchAll(pattern)) {
      const id = match[1];
      const xml = match[0];
      if (id && xml) {
        entries[id] = xml;
      }
    }
  }
  return entries;
}

function parseOrderedPathGeometry(rawShapeXml: string): PathGeometry {
  const pathMatches = [...rawShapeXml.matchAll(/<a:path\b([^>]*)>([\s\S]*?)<\/a:path>/g)];
  if (!pathMatches.length) {
    return {};
  }

  const commands: PresentationPathCommand[] = [];
  const commandPattern = /<(a:moveTo|a:lnTo|a:cubicBezTo|a:close)\b[^>]*>([\s\S]*?)<\/\1>|<a:close\b[^>]*\/>/g;
  let viewportWidth = 0;
  let viewportHeight = 0;
  for (const pathMatch of pathMatches) {
    viewportWidth = Math.max(viewportWidth, Number((pathMatch[1].match(/\bw="([^"]+)"/)?.[1]) ?? 0));
    viewportHeight = Math.max(viewportHeight, Number((pathMatch[1].match(/\bh="([^"]+)"/)?.[1]) ?? 0));
    for (const match of pathMatch[0].matchAll(commandPattern)) {
      const type = match[1] ?? 'a:close';
      const body = match[2] ?? '';
      if (type === 'a:close') {
        commands.push({ type: 'close' });
        continue;
      }
      const points = [...body.matchAll(/<a:pt\b[^>]*x="([^"]+)"[^>]*y="([^"]+)"[^>]*\/>/g)].map((point) => ({
        x: Number(point[1]),
        y: Number(point[2])
      }));
      if (type === 'a:moveTo' || type === 'a:lnTo') {
        const point = points[0];
        if (point && Number.isFinite(point.x) && Number.isFinite(point.y)) {
          commands.push({ type: type === 'a:moveTo' ? 'moveTo' : 'lineTo', x: point.x, y: point.y });
        }
        continue;
      }
      if (type === 'a:cubicBezTo' && points.length >= 3) {
        commands.push({
          type: 'cubicTo',
          x1: points[0]?.x,
          y1: points[0]?.y,
          x2: points[1]?.x,
          y2: points[1]?.y,
          x: points[2]?.x,
          y: points[2]?.y
        });
      }
    }
  }

  return {
    commands: commands.length ? commands : undefined,
    viewport: viewportWidth > 0 && viewportHeight > 0
      ? { width: viewportWidth, height: viewportHeight }
      : undefined
  };
}

function parseTextStyle(shape: Record<string, unknown>, theme?: PresentationTheme): PresentationTextStyle | undefined {
  const textBody = xmlChild<Record<string, unknown>>(shape, 'p:txBody');
  if (!textBody) {
    return undefined;
  }

  const paragraph = xmlChild<Record<string, unknown>>(textBody, 'a:p');
  const paragraphProperties = xmlChild<Record<string, unknown>>(paragraph, 'a:pPr');
  const run = xmlChild<Record<string, unknown>>(paragraph, 'a:r');
  const runProperties = xmlChild<Record<string, unknown>>(run, 'a:rPr');
  const endParagraphProperties = xmlChild<Record<string, unknown>>(paragraph, 'a:endParaRPr');
  const solidFill = xmlChild<Record<string, unknown>>(runProperties, 'a:solidFill')
    ?? xmlChild<Record<string, unknown>>(endParagraphProperties, 'a:solidFill');
  const latin = xmlChild<Record<string, unknown>>(runProperties, 'a:latin')
    ?? xmlChild<Record<string, unknown>>(endParagraphProperties, 'a:latin');
  const eastAsian = xmlChild<Record<string, unknown>>(runProperties, 'a:ea')
    ?? xmlChild<Record<string, unknown>>(endParagraphProperties, 'a:ea');
  const complexScript = xmlChild<Record<string, unknown>>(runProperties, 'a:cs')
    ?? xmlChild<Record<string, unknown>>(endParagraphProperties, 'a:cs');
  const color = solidFill ? resolveColor(solidFill, theme)?.color : undefined;
  const size = xmlAttr(runProperties, 'sz') ?? xmlAttr(endParagraphProperties, 'sz');

  if (!runProperties && !endParagraphProperties && !paragraphProperties) {
    return undefined;
  }

  return {
    color,
    fontSizePt: size ? Number(size) / 100 : undefined,
    fontFamily: xmlAttr(latin, 'typeface') ?? xmlAttr(eastAsian, 'typeface') ?? xmlAttr(complexScript, 'typeface') ?? undefined,
    bold: xmlAttr(runProperties, 'b') === '1' || xmlAttr(endParagraphProperties, 'b') === '1'
      ? true
      : xmlAttr(runProperties, 'b') === '0' || xmlAttr(endParagraphProperties, 'b') === '0'
        ? false
        : undefined,
    italic: xmlAttr(runProperties, 'i') === '1' || xmlAttr(endParagraphProperties, 'i') === '1'
      ? true
      : xmlAttr(runProperties, 'i') === '0' || xmlAttr(endParagraphProperties, 'i') === '0'
        ? false
        : undefined,
    align: xmlAttr(paragraphProperties, 'algn') ?? undefined
  };
}

function parseLayoutInfo(
  graph: PackageGraph,
  layoutUri: string,
  themeCache: Record<string, PresentationTheme>
): {
  layoutUri: string;
  layoutName?: string;
  masterUri?: string;
  masterName?: string;
  themeUri?: string;
  layoutBackground?: PresentationFill;
  masterBackground?: PresentationFill;
  layoutShapes: SlideShape[];
  masterShapes: SlideShape[];
} {
  const layoutXml = getParsedXmlPart(graph, layoutUri);
  const layoutRoot = layoutXml?.document['p:sldLayout'] as Record<string, unknown> | undefined;
  const layoutCommonSlideData = xmlChild<Record<string, unknown>>(layoutRoot, 'p:cSld');
  const layoutName = layoutRoot ? xmlAttr(layoutRoot, 'matchingName') ?? xmlAttr(layoutRoot, 'type') : undefined;
  const masterRelationship = relationshipsFor(graph, layoutUri).find((relationship) => relationship.type.includes('/slideMaster'));
  const masterUri = masterRelationship?.resolvedTarget ?? undefined;
  let masterName: string | undefined;
  let themeUri: string | undefined;
  let masterBackground: PresentationFill | undefined;
  let masterShapes: SlideShape[] = [];

  if (masterUri) {
    const masterXml = getParsedXmlPart(graph, masterUri);
    const masterRoot = masterXml?.document['p:sldMaster'] as Record<string, unknown> | undefined;
    const masterCommonSlideData = xmlChild<Record<string, unknown>>(masterRoot, 'p:cSld');
    masterName = masterRoot ? xmlAttr(masterRoot, 'preserve') ?? 'slide-master' : 'slide-master';
    const themeRelationship = relationshipsFor(graph, masterUri).find((relationship) => relationship.type.includes('/theme'));
    themeUri = themeRelationship?.resolvedTarget ?? undefined;
    if (themeUri && !themeCache[themeUri]) {
      themeCache[themeUri] = parseThemeInfo(graph, themeUri);
    }
    const theme = themeUri ? themeCache[themeUri] : undefined;
    masterBackground = parseBackground(masterCommonSlideData, relationshipsFor(graph, masterUri), theme);
    masterShapes = collectShapes(xmlChild<Record<string, unknown>>(masterCommonSlideData, 'p:spTree'), relationshipsFor(graph, masterUri), theme);
  }

  const theme = themeUri ? themeCache[themeUri] : undefined;
  const layoutBackground = parseBackground(layoutCommonSlideData, relationshipsFor(graph, layoutUri), theme);
  const layoutShapes = collectShapes(xmlChild<Record<string, unknown>>(layoutCommonSlideData, 'p:spTree'), relationshipsFor(graph, layoutUri), theme);

  return { layoutUri, layoutName, masterUri, masterName, themeUri, layoutBackground, masterBackground, layoutShapes, masterShapes };
}

function applyInheritedPlaceholderProperties(slideShapes: SlideShape[], inheritedShapes: SlideShape[]): SlideShape[] {
  return slideShapes.map((shape) => {
    if (!shape.placeholderType && !shape.placeholderIndex) {
      return shape;
    }

    const match = inheritedShapes.find((candidate) =>
      candidate.placeholderType === shape.placeholderType
      && candidate.placeholderIndex === shape.placeholderIndex
    );
    if (!match) {
      return shape;
    }

    return {
      ...match,
      ...shape,
      transform: shape.transform ?? match.transform,
      fill: shape.fill ?? match.fill,
      line: shape.line ?? match.line,
      textStyle: shape.textStyle ?? match.textStyle,
      shapeType: shape.shapeType ?? match.shapeType
    };
  });
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
  const colors = Object.fromEntries(
    ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
      .flatMap((key) => {
        const colorNode = xmlChild<Record<string, unknown>>(colorScheme, `a:${key}`);
        const resolved = colorNode ? resolveColor(colorNode, undefined) : undefined;
        return resolved?.color ? [[key, resolved.color]] : [];
      })
  );

  return {
    uri: themeUri,
    name: theme ? xmlAttr(theme, 'name') : undefined,
    colorSchemeName: colorScheme ? xmlAttr(colorScheme, 'name') : undefined,
    majorLatinFont: majorLatin ? xmlAttr(majorLatin, 'typeface') : undefined,
    minorLatinFont: minorLatin ? xmlAttr(minorLatin, 'typeface') : undefined,
    colors
  };
}

function parseBackground(
  commonSlideData: Record<string, unknown> | undefined,
  relationships: ReturnType<typeof relationshipsFor>,
  theme?: PresentationTheme
): PresentationFill | undefined {
  const background = xmlChild<Record<string, unknown>>(commonSlideData, 'p:bg');
  const backgroundProperties = xmlChild<Record<string, unknown>>(background, 'p:bgPr');
  const backgroundReference = xmlChild<Record<string, unknown>>(background, 'p:bgRef');
  return parseFill(backgroundProperties ?? backgroundReference, relationships, theme);
}

function parseFill(
  shapeProperties: Record<string, unknown> | undefined,
  relationships: ReturnType<typeof relationshipsFor>,
  theme?: PresentationTheme,
  styleNode?: Record<string, unknown>,
  shapeType?: string
): PresentationFill | undefined {
  if (!shapeProperties) {
    return undefined;
  }

  if (xmlChild<Record<string, unknown> | string>(shapeProperties, 'a:noFill') !== undefined) {
    return { kind: 'none' };
  }

  const solidFill = xmlChild<Record<string, unknown>>(shapeProperties, 'a:solidFill');
  if (solidFill) {
    const resolved = resolveColor(solidFill, theme);
    return {
      kind: 'solid',
      color: resolved?.color,
      opacity: resolved?.opacity
    };
  }

  const gradientFill = xmlChild<Record<string, unknown>>(shapeProperties, 'a:gradFill');
  if (gradientFill) {
    return {
      kind: 'gradient',
      gradientStops: xmlChildren<Record<string, unknown>>(xmlChild<Record<string, unknown>>(gradientFill, 'a:gsLst'), 'a:gs')
        .map((stop) => ({
          position: Number(xmlAttr(stop, 'pos') ?? 0) / 1000,
          ...resolveColor(stop, theme)
        }))
        .filter((stop) => stop.color),
      angleDeg: (() => {
        const linear = xmlChild<Record<string, unknown>>(gradientFill, 'a:lin');
        const value = xmlAttr(linear, 'ang');
        return value ? Number(value) / 60_000 : undefined;
      })()
    };
  }

  const blipFill = xmlChild<Record<string, unknown>>(shapeProperties, 'a:blipFill');
  const blip = xmlChild<Record<string, unknown>>(blipFill, 'a:blip');
  const relationshipId = xmlAttr(blip, 'r:embed');
  if (relationshipId) {
    return {
      kind: 'image',
      targetUri: relationships.find((relationship) => relationship.id === relationshipId)?.resolvedTarget ?? undefined
    };
  }

  const directColor = resolveColor(shapeProperties, theme);
  if (directColor?.color) {
    return {
      kind: 'solid',
      color: directColor.color,
      opacity: directColor.opacity
    };
  }

  const styleFill = xmlChild<Record<string, unknown>>(styleNode, 'a:fillRef');
  if (shapeType === 'custom' && xmlChild<Record<string, unknown>>(shapeProperties, 'a:ln')) {
    return undefined;
  }
  const styleFillColor = resolveColor(styleFill, theme);
  if (styleFillColor?.color) {
    return {
      kind: 'solid',
      color: styleFillColor.color,
      opacity: styleFillColor.opacity
    };
  }

  return undefined;
}

function parseLine(shapeProperties: Record<string, unknown> | undefined, theme?: PresentationTheme, styleNode?: Record<string, unknown>): PresentationLine | undefined {
  const line = xmlChild<Record<string, unknown>>(shapeProperties, 'a:ln');
  if (!line) {
    return undefined;
  }

  if (xmlChild<Record<string, unknown> | string>(line, 'a:noFill') !== undefined) {
    return {
      kind: 'none',
      width: (() => { const value = xmlAttr(line, 'w'); return value ? Number(value) : undefined; })()
    };
  }

  const solidFill = xmlChild<Record<string, unknown>>(line, 'a:solidFill');
  const resolved = solidFill ? resolveColor(solidFill, theme) : undefined;
  const styleLine = xmlChild<Record<string, unknown>>(styleNode, 'a:lnRef');
  const styleResolved = styleLine ? resolveColor(styleLine, theme) : undefined;
  return {
    kind: 'solid',
    color: resolved?.color ?? styleResolved?.color,
    opacity: resolved?.opacity ?? styleResolved?.opacity,
    width: (() => { const value = xmlAttr(line, 'w'); return value ? Number(value) : xmlAttr(styleLine, 'idx') ? Number(xmlAttr(styleLine, 'idx')) * 6350 : undefined; })()
  };
}

function resolveColor(parent: Record<string, unknown> | undefined, theme?: PresentationTheme): { color?: string; opacity?: number } | undefined {
  if (!parent) {
    return undefined;
  }

  const srgb = xmlChild<Record<string, unknown>>(parent, 'a:srgbClr');
  if (srgb) {
    return {
      color: applyColorTransforms(`#{val}`.replace('{val}', xmlAttr(srgb, 'val') ?? '').replace(/^#$/, ''), srgb),
      opacity: parseAlpha(srgb)
    };
  }

  const scheme = xmlChild<Record<string, unknown>>(parent, 'a:schemeClr');
  if (scheme) {
    const schemeValue = xmlAttr(scheme, 'val') ?? '';
    return {
      color: applyColorTransforms(theme?.colors?.[schemeValue] ?? defaultThemeColors[schemeValue], scheme),
      opacity: parseAlpha(scheme)
    };
  }

  const sys = xmlChild<Record<string, unknown>>(parent, 'a:sysClr');
  if (sys) {
    return {
      color: applyColorTransforms(`#{val}`.replace('{val}', xmlAttr(sys, 'lastClr') ?? '').replace(/^#$/, ''), sys),
      opacity: parseAlpha(sys)
    };
  }

  return undefined;
}

function applyColorTransforms(color: string | undefined, node: Record<string, unknown>): string | undefined {
  if (!color || !/^#[0-9A-Fa-f]{6}$/.test(color)) {
    return color;
  }

  let [r, g, b] = [color.slice(1, 3), color.slice(3, 5), color.slice(5, 7)].map((value) => Number.parseInt(value, 16));
  const lumMod = Number(xmlAttr(xmlChild<Record<string, unknown>>(node, 'a:lumMod'), 'val') ?? 100000) / 100000;
  const lumOff = Number(xmlAttr(xmlChild<Record<string, unknown>>(node, 'a:lumOff'), 'val') ?? 0) / 100000;
  const shade = Number(xmlAttr(xmlChild<Record<string, unknown>>(node, 'a:shade'), 'val') ?? 100000) / 100000;
  const tint = Number(xmlAttr(xmlChild<Record<string, unknown>>(node, 'a:tint'), 'val') ?? 0) / 100000;

  const transformChannel = (channel: number): number => {
    let value = channel * lumMod + 255 * lumOff;
    value *= shade;
    value = value + (255 - value) * tint;
    return Math.max(0, Math.min(255, Math.round(value)));
  };

  r = transformChannel(r);
  g = transformChannel(g);
  b = transformChannel(b);
  return `#${[r, g, b].map((value) => value.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function parseAlpha(node: Record<string, unknown>): number | undefined {
  const alpha = xmlChild<Record<string, unknown>>(node, 'a:alpha');
  const value = xmlAttr(alpha, 'val');
  return value ? Number(value) / 100000 : undefined;
}

const defaultThemeColors: Record<string, string> = {
  dk1: '#000000',
  lt1: '#FFFFFF',
  dk2: '#1F497D',
  lt2: '#EEECE1',
  accent1: '#4F81BD',
  accent2: '#C0504D',
  accent3: '#9BBB59',
  accent4: '#8064A2',
  accent5: '#4BACC6',
  accent6: '#F79646',
  hlink: '#0000FF',
  folHlink: '#800080',
  bg1: '#FFFFFF',
  tx1: '#000000',
  bg2: '#EEECE1',
  tx2: '#1F497D',
  phClr: '#4F81BD'
};

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
  for (const nodeType of ['p:par', 'p:seq', 'p:anim', 'p:animClr', 'p:animMotion', 'p:set', 'p:cmd']) {
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
        targetShapeId: xmlAttr(shapeTarget, 'spid') ?? undefined,
        colorSpace: xmlAttr(node, 'clrSpc') ?? undefined,
        colorDirection: xmlAttr(node, 'dir') ?? undefined,
        motionOrigin: xmlAttr(node, 'origin') ?? undefined,
        motionPath: xmlAttr(node, 'path') ?? undefined,
        motionPathEditMode: xmlAttr(node, 'pathEditMode') ?? undefined,
        commandName: xmlAttr(node, 'cmd') ?? undefined,
        commandType: xmlAttr(node, 'type') ?? undefined
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
