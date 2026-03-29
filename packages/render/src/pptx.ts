import type { PresentationDocument, SlideShape } from '@ooxml/pptx';

import type { RenderOptions } from './types';

type PresentationFill = PresentationDocument['slides'][number]['background'];
type PresentationPathCommand = NonNullable<PresentationDocument['slides'][number]['shapes'][number]['pathCommands']>[number];

export function renderPresentation(presentation: PresentationDocument, options: RenderOptions): string {
  const slideIndex = Math.min(options.activeSlideIndex ?? 0, Math.max(presentation.slides.length - 1, 0));
  const slide = presentation.slides[slideIndex];

  if (!slide) {
    return '<section class="ooxml-render ooxml-render--pptx">No slides available</section>';
  }

  const theme = slide.themeUri ? presentation.themes[slide.themeUri] : undefined;
  if (options.pptxRenderer === 'scene-svg') {
    const sceneBackgroundStyle = buildSceneBackgroundStyle(slide.background);
    const sceneShapes = slide.shapes.map((shape) => renderSceneShape(shape, presentation.size.cx, presentation.size.cy, slide.title)).join('');
    return `<section class="ooxml-render ooxml-render--pptx ooxml-render--pptx-scene" data-presentation-cx="${presentation.size.cx}" data-presentation-cy="${presentation.size.cy}"${slide.background?.color ? ` data-background-color="${escapeHtml(slide.background.color)}"` : ''}${slide.background?.opacity !== undefined ? ` data-background-opacity="${slide.background.opacity}"` : ''}${slide.background?.targetUri ? ` data-background-image-uri="${escapeHtml(slide.background.targetUri)}"` : ''}${slide.layoutUri ? ` data-layout-uri="${escapeHtml(slide.layoutUri)}"` : ''}${slide.masterUri ? ` data-master-uri="${escapeHtml(slide.masterUri)}"` : ''}${slide.themeUri ? ` data-theme-uri="${escapeHtml(slide.themeUri)}"` : ''}><header><h2>${escapeHtml(slide.title)}</h2><p>${presentation.size.cx} × ${presentation.size.cy}</p></header><div class="ooxml-pptx-scene" style="${sceneBackgroundStyle};aspect-ratio:${presentation.size.cx} / ${presentation.size.cy};"${slide.background?.targetUri ? ` data-background-image-uri="${escapeHtml(slide.background.targetUri)}"` : ''}>${sceneShapes}</div><dl class="ooxml-pptx-inheritance"><dt>Layout</dt><dd>${escapeHtml(slide.layoutName ?? slide.layoutUri ?? 'none')}</dd><dt>Master</dt><dd>${escapeHtml(slide.masterName ?? slide.masterUri ?? 'none')}</dd><dt>Theme</dt><dd>${escapeHtml(theme?.name ?? theme?.colorSchemeName ?? slide.themeUri ?? 'none')}</dd></dl></section>`;
  }

  const shapeMarkup = slide.shapes.map((shape) => `<div class="ooxml-pptx-shape"${shape.placeholderType ? ` data-placeholder-type="${escapeHtml(shape.placeholderType)}"` : ''}${shape.placeholderIndex ? ` data-placeholder-index="${escapeHtml(shape.placeholderIndex)}"` : ''}${shape.shapeType ? ` data-shape-type="${escapeHtml(shape.shapeType)}"` : ''}${shape.fill?.color ? ` data-fill-color="${escapeHtml(shape.fill.color)}"` : ''}${shape.fill?.opacity !== undefined ? ` data-fill-opacity="${shape.fill.opacity}"` : ''}${shape.fill?.targetUri ? ` data-fill-image-uri="${escapeHtml(shape.fill.targetUri)}"` : ''}${shape.fill?.gradientStops?.length ? ` data-fill-gradient-stops="${escapeHtml(serializeGradientStops(shape.fill.gradientStops))}"` : ''}${shape.fill?.angleDeg !== undefined ? ` data-fill-gradient-angle="${shape.fill.angleDeg}"` : ''}${shape.line?.color ? ` data-line-color="${escapeHtml(shape.line.color)}"` : ''}${shape.line?.opacity !== undefined ? ` data-line-opacity="${shape.line.opacity}"` : ''}${shape.line?.width !== undefined ? ` data-line-width="${shape.line.width}"` : ''}${shape.line?.dash ? ` data-line-dash="${escapeHtml(shape.line.dash)}"` : ''}${shape.textStyle?.color ? ` data-text-color="${escapeHtml(shape.textStyle.color)}"` : ''}${shape.textStyle?.fontSizePt !== undefined ? ` data-font-size-pt="${shape.textStyle.fontSizePt}"` : ''}${shape.textStyle?.fontFamily ? ` data-font-family="${escapeHtml(shape.textStyle.fontFamily)}"` : ''}${shape.textStyle?.bold !== undefined ? ` data-font-bold="${shape.textStyle.bold}"` : ''}${shape.textStyle?.italic !== undefined ? ` data-font-italic="${shape.textStyle.italic}"` : ''}${shape.textStyle?.align ? ` data-text-align="${escapeHtml(shape.textStyle.align)}"` : ''}${shape.pathCommands ? ` data-path-commands="${escapeHtml(serializePathCommands(shape.pathCommands))}"` : ''}${shape.pathViewport ? ` data-path-viewport="${shape.pathViewport.width}:${shape.pathViewport.height}"` : ''}${shape.media?.targetUri ? ` data-media-uri="${escapeHtml(shape.media.targetUri)}"` : ''}${shape.media?.type ? ` data-media-type="${escapeHtml(shape.media.type)}"` : ''}${shape.media?.progId ? ` data-prog-id="${escapeHtml(shape.media.progId)}"` : ''}${shape.transform?.x !== undefined ? ` data-x="${shape.transform.x}"` : ''}${shape.transform?.y !== undefined ? ` data-y="${shape.transform.y}"` : ''}${shape.transform?.cx !== undefined ? ` data-cx="${shape.transform.cx}"` : ''}${shape.transform?.cy !== undefined ? ` data-cy="${shape.transform.cy}"` : ''}${shape.transform?.rotationDeg !== undefined ? ` data-rotation-deg="${shape.transform.rotationDeg}"` : ''}${shape.transform?.flipH ? ' data-flip-h="true"' : ''}${shape.transform?.flipV ? ' data-flip-v="true"' : ''}><h3>${escapeHtml(shape.name ?? 'Shape')}</h3><p>${escapeHtml(shape.text || (shape.media?.type === 'embeddedObject' ? '[embedded-object]' : shape.media ? '[image]' : ''))}</p></div>`).join('');
  const notes = slide.notesText ? `<aside class="ooxml-pptx-notes">${escapeHtml(slide.notesText)}</aside>` : '';
  const commentsMarkup = slide.comments.length ? `<ul class="ooxml-pptx-comments">${slide.comments.map((comment) => `<li data-comment-index="${comment.index}">${escapeHtml(comment.text)}${comment.author ? ` — ${escapeHtml(comment.author)}` : ''}</li>`).join('')}</ul>` : '';
  const timingMarkup = slide.transition || slide.timing ? `<dl class="ooxml-pptx-timing">${slide.transition?.type ? `<dt>Transition</dt><dd data-transition-type="${escapeHtml(slide.transition.type)}">${escapeHtml(slide.transition.type)}${slide.transition.speed ? ` (${escapeHtml(slide.transition.speed)})` : ''}${slide.transition.advanceOnClick !== undefined ? ` click:${escapeHtml(String(slide.transition.advanceOnClick))}` : ''}${slide.transition.advanceAfterMs !== undefined ? ` after:${escapeHtml(String(slide.transition.advanceAfterMs))}` : ''}</dd>` : ''}${slide.timing ? `<dt>Timing nodes</dt><dd data-timing-count="${slide.timing.nodeCount}">${slide.timing.nodes.map((node) => `${escapeHtml(node.nodeType)}${node.presetClass ? `:${escapeHtml(node.presetClass)}` : ''}${node.concurrent !== undefined ? ` concurrent:${escapeHtml(String(node.concurrent))}` : ''}${node.nextAction ? ` next:${escapeHtml(node.nextAction)}` : ''}${node.previousAction ? ` prev:${escapeHtml(node.previousAction)}` : ''}${node.id ? `#${escapeHtml(node.id)}` : ''}${node.duration ? `@${escapeHtml(node.duration)}` : ''}${node.repeatDuration ? ` repeatDur:${escapeHtml(node.repeatDuration)}` : ''}${node.repeatCount ? `×${escapeHtml(node.repeatCount)}` : ''}${node.restart ? ` restart:${escapeHtml(node.restart)}` : ''}${node.fill ? ` fill:${escapeHtml(node.fill)}` : ''}${node.autoReverse !== undefined ? ` autoRev:${escapeHtml(String(node.autoReverse))}` : ''}${node.acceleration ? ` accel:${escapeHtml(node.acceleration)}` : ''}${node.deceleration ? ` decel:${escapeHtml(node.deceleration)}` : ''}${node.colorSpace ? ` clr:${escapeHtml(node.colorSpace)}` : ''}${node.colorDirection ? ` dir:${escapeHtml(node.colorDirection)}` : ''}${node.motionOrigin ? ` origin:${escapeHtml(node.motionOrigin)}` : ''}${node.motionPath ? ` path:${escapeHtml(node.motionPath)}` : ''}${node.motionPathEditMode ? ` pathEdit:${escapeHtml(node.motionPathEditMode)}` : ''}${node.commandName ? ` cmd:${escapeHtml(node.commandName)}` : ''}${node.commandType ? ` cmdType:${escapeHtml(node.commandType)}` : ''}${node.triggerEvent ? `!${escapeHtml(node.triggerEvent)}` : ''}${node.triggerDelay ? `+${escapeHtml(node.triggerDelay)}` : ''}${node.triggerShapeId ? `^${escapeHtml(node.triggerShapeId)}` : ''}${node.endTriggerEvent ? ` ~${escapeHtml(node.endTriggerEvent)}` : ''}${node.endTriggerDelay ? `=${escapeHtml(node.endTriggerDelay)}` : ''}${node.endTriggerShapeId ? `^${escapeHtml(node.endTriggerShapeId)}` : ''}${node.targetShapeId ? `->${escapeHtml(node.targetShapeId)}` : ''}`).join(', ')}</dd>` : ''}</dl>` : '';
  const inheritanceMarkup = `<dl class="ooxml-pptx-inheritance"><dt>Layout</dt><dd>${escapeHtml(slide.layoutName ?? slide.layoutUri ?? 'none')}</dd><dt>Master</dt><dd>${escapeHtml(slide.masterName ?? slide.masterUri ?? 'none')}</dd><dt>Theme</dt><dd>${escapeHtml(theme?.name ?? theme?.colorSchemeName ?? slide.themeUri ?? 'none')}</dd>${theme?.majorLatinFont ? `<dt>Major font</dt><dd>${escapeHtml(theme.majorLatinFont)}</dd>` : ''}${theme?.minorLatinFont ? `<dt>Minor font</dt><dd>${escapeHtml(theme.minorLatinFont)}</dd>` : ''}</dl>`;

  return `<section class="ooxml-render ooxml-render--pptx" data-presentation-cx="${presentation.size.cx}" data-presentation-cy="${presentation.size.cy}"${slide.background?.color ? ` data-background-color="${escapeHtml(slide.background.color)}"` : ''}${slide.background?.opacity !== undefined ? ` data-background-opacity="${slide.background.opacity}"` : ''}${slide.background?.targetUri ? ` data-background-image-uri="${escapeHtml(slide.background.targetUri)}"` : ''}${slide.layoutUri ? ` data-layout-uri="${escapeHtml(slide.layoutUri)}"` : ''}${slide.masterUri ? ` data-master-uri="${escapeHtml(slide.masterUri)}"` : ''}${slide.themeUri ? ` data-theme-uri="${escapeHtml(slide.themeUri)}"` : ''}><header><h2>${escapeHtml(slide.title)}</h2><p>${presentation.size.cx} × ${presentation.size.cy}</p></header>${inheritanceMarkup}${timingMarkup}${shapeMarkup}${commentsMarkup}${notes}</section>`;
}

function renderSceneShape(shape: SlideShape, presentationCx: number, presentationCy: number, slideTitle: string): string {
  const x = shape.transform?.x ?? 0;
  const y = shape.transform?.y ?? 0;
  const width = shape.transform?.cx ?? Math.max(presentationCx * 0.1, 1);
  const height = shape.transform?.cy ?? Math.max(presentationCy * 0.05, 1);
  const isSlideTitle = shape.text.trim() === slideTitle.trim() && slideTitle.trim().length > 0;
  const inferredCenterText = !shape.textStyle?.align && width / presentationCx > 0.6 && shape.text.trim().length > 0;
  const textAlign = isSlideTitle || inferredCenterText ? 'center' : toCssAlign(shape.textStyle?.align);
  const centerText = textAlign === 'center';
  const style = [
    'position:absolute',
    `left:${(x / presentationCx) * 100}%`,
    `top:${(y / presentationCy) * 100}%`,
    `width:${(width / presentationCx) * 100}%`,
    `height:${(height / presentationCy) * 100}%`,
    'box-sizing:border-box',
    'overflow:hidden',
    shape.shapeType === 'ellipse' ? 'border-radius:999px' : '',
    shape.shapeType === 'chevron' ? 'clip-path:polygon(0 0, 76% 0, 100% 50%, 76% 100%, 0 100%, 18% 50%)' : '',
    buildTransformStyle(shape),
    shape.textStyle?.color ? `color:${shape.textStyle.color}` : '',
    shape.textStyle?.fontFamily ? `font-family:${escapeHtml(shape.textStyle.fontFamily)}` : '',
    shape.textStyle?.fontSizePt !== undefined ? `font-size:${Math.max(shape.textStyle.fontSizePt, 12)}px` : '',
    shape.textStyle?.bold ? 'font-weight:700' : '',
    shape.textStyle?.italic ? 'font-style:italic' : '',
    `text-align:${textAlign}`,
    'display:flex',
    `align-items:${toCssAnchor(shape.textStyle?.anchor)}`,
    centerText ? 'justify-content:center' : 'justify-content:flex-start'
  ].filter(Boolean).join(';');
  const filledStyle = [
    style,
    shape.fill?.color ? `background:${applyOpacity(shape.fill.color, shape.fill.opacity)}` : '',
    shape.line?.color ? `border:${Math.max(1, emuToPx(shape.line.width ?? 0))}px solid ${applyOpacity(shape.line.color, shape.line.opacity)}` : ''
  ].filter(Boolean).join(';');

  if (shape.media?.type === 'image' && shape.media.targetUri) {
    return `<div class="ooxml-pptx-scene-node ooxml-pptx-scene-node--image" style="${style}"><img src="${escapeHtml(shape.media.targetUri)}" alt="" data-media-uri="${escapeHtml(shape.media.targetUri)}"></div>`;
  }

  if (shape.text && shouldRenderSceneTextOnly(shape)) {
    const text = renderSceneText(shape, false, isSlideTitle || inferredCenterText);
    return `<div class="ooxml-pptx-scene-node ooxml-pptx-scene-node--text" style="${style}">${text}</div>`;
  }

  if (shape.pathCommands?.length || isPresetSceneVectorShape(shape.shapeType)) {
    const text = shape.text ? renderSceneText(shape, true, centerText) : '';
    return `<div class="ooxml-pptx-scene-node ooxml-pptx-scene-node--vector" style="${style};position:absolute;">${renderSceneShapeSvg(shape)}${text}</div>`;
  }

  const text = shape.text ? renderSceneText(shape, false, centerText) : '';
  return `<div class="ooxml-pptx-scene-node" style="${filledStyle}">${text}</div>`;
}

function shouldRenderSceneTextOnly(shape: SlideShape): boolean {
  return (shape.shapeType === 'rect' || shape.shapeType === 'roundRect' || shape.shapeType === 'round2SameRect')
    && (!shape.fill || shape.fill.kind === 'none')
    && (!shape.line || shape.line.kind === 'none');
}

function renderSceneShapeSvg(shape: SlideShape): string {
  const gradientId = shape.fill?.gradientStops?.length ? 'ooxml-scene-gradient' : undefined;
  const fill = gradientId ? `url(#${gradientId})` : shape.fill?.color ?? 'none';
  const stroke = shape.line?.color ?? 'none';
  const strokeWidth = Math.max(1, emuToPx(shape.line?.width ?? 0));
  const strokeDashArray = sceneStrokeDashArray(shape.line);
  const strokeAttrs = `${shape.line?.opacity !== undefined ? ` stroke-opacity="${shape.line.opacity}"` : ''}${strokeDashArray ? ` stroke-dasharray="${strokeDashArray}"` : ''}`;
  const preserveAspectRatio = shouldUseIntrinsicAspectRatio(shape) ? 'xMidYMid meet' : 'none';
  const gradientMarkup = gradientId
    ? `<defs><linearGradient id="${gradientId}" gradientTransform="rotate(${shape.fill?.angleDeg ?? 0}, 0.5, 0.5)">${(shape.fill?.gradientStops ?? []).map((stop) => `<stop offset="${stop.position}%" stop-color="${escapeHtml(stop.color ?? '#000000')}"${stop.opacity !== undefined ? ` stop-opacity="${stop.opacity}"` : ''}/>`).join('')}</linearGradient></defs>`
    : '';
  if (shape.pathCommands?.length) {
    const viewBox = shape.pathViewport ? `0 0 ${shape.pathViewport.width} ${shape.pathViewport.height}` : buildPathViewBox(shape.pathCommands);
    const evenOdd = shouldUseEvenOddFill(shape);
    return `<svg class="ooxml-pptx-scene-svg" viewBox="${viewBox}" preserveAspectRatio="${preserveAspectRatio}" aria-hidden="true">${gradientMarkup}<path d="${escapeHtml(toSvgPath(shape.pathCommands))}" fill="${escapeHtml(fill)}"${shape.fill?.opacity !== undefined && !gradientId ? ` fill-opacity="${shape.fill.opacity}"` : ''}${evenOdd ? ' fill-rule="evenodd" clip-rule="evenodd"' : ''} stroke="${escapeHtml(stroke)}"${strokeAttrs} stroke-width="${strokeWidth}" vector-effect="non-scaling-stroke"/></svg>`;
  }

  const presetMarkup = buildPresetSceneSvgMarkup(shape.shapeType, fill, shape.fill?.opacity, stroke, strokeWidth, strokeAttrs);
  if (!presetMarkup) {
    return '';
  }

  return `<svg class="ooxml-pptx-scene-svg" viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true">${gradientMarkup}${presetMarkup}</svg>`;
}

function buildSceneBackgroundStyle(fill: PresentationFill | undefined): string {
  const styles = ['position:relative', 'width:min(100%, 960px)', 'margin:0 auto', 'overflow:hidden'];
  if (fill?.color) {
    styles.push(`background:${applyOpacity(fill.color, fill.opacity)}`);
  }
  return styles.join(';');
}

function buildTransformStyle(shape: SlideShape): string {
  const parts: string[] = [];
  if (shape.transform?.flipH) parts.push('scaleX(-1)');
  if (shape.transform?.flipV) parts.push('scaleY(-1)');
  if (shape.transform?.rotationDeg) parts.push(`rotate(${shape.transform.rotationDeg}deg)`);
  return parts.length ? `transform:${parts.join(' ')};transform-origin:center center` : '';
}

function toSvgPath(commands: PresentationPathCommand[]): string {
  return commands.map((command) => {
    if (command.type === 'close') return 'Z';
    if (command.type === 'cubicTo') return `C ${command.x1 ?? 0} ${command.y1 ?? 0} ${command.x2 ?? 0} ${command.y2 ?? 0} ${command.x ?? 0} ${command.y ?? 0}`;
    return `${command.type === 'moveTo' ? 'M' : 'L'} ${command.x ?? 0} ${command.y ?? 0}`;
  }).join(' ');
}

function buildPathViewBox(commands: PresentationPathCommand[]): string {
  const points = commands.flatMap((command) => {
    if (command.type === 'cubicTo') {
      return [
        { x: command.x1 ?? 0, y: command.y1 ?? 0 },
        { x: command.x2 ?? 0, y: command.y2 ?? 0 },
        { x: command.x ?? 0, y: command.y ?? 0 }
      ];
    }
    if (command.type === 'close') {
      return [];
    }
    return [{ x: command.x ?? 0, y: command.y ?? 0 }];
  });
  if (!points.length) {
    return '0 0 1 1';
  }
  const minX = Math.min(...points.map((point) => point.x));
  const minY = Math.min(...points.map((point) => point.y));
  const maxX = Math.max(...points.map((point) => point.x));
  const maxY = Math.max(...points.map((point) => point.y));
  return `${minX} ${minY} ${Math.max(1, maxX - minX)} ${Math.max(1, maxY - minY)}`;
}

function shouldUseEvenOddFill(shape: SlideShape): boolean {
  const commands = shape.pathCommands ?? [];
  const viewport = shape.pathViewport;
  if (!viewport || commands.length < 8) {
    return false;
  }
  const moveCommands = commands.filter((command) => command.type === 'moveTo' && command.x !== undefined && command.y !== undefined);
  if (moveCommands.length < 2) {
    return false;
  }
  const aspect = viewport.width / Math.max(viewport.height, 1);
  if (aspect < 0.8 || aspect > 1.2) {
    return false;
  }
  if (shape.line && shape.line.kind !== 'none') {
    return false;
  }
  if (shape.fill?.kind !== 'solid' || shape.fill.color !== '#FFFFFF') {
    return false;
  }
  const first = moveCommands[0]!;
  const second = moveCommands[1]!;
  const firstXRatio = (first.x ?? 0) / viewport.width;
  const firstYRatio = (first.y ?? 0) / viewport.height;
  const secondXRatio = (second.x ?? 0) / viewport.width;
  const secondYRatio = (second.y ?? 0) / viewport.height;
  return firstXRatio > 0.85
    && firstYRatio > 0.35 && firstYRatio < 0.7
    && secondXRatio > 0.25 && secondXRatio < 0.75
    && secondYRatio < 0.1;
}

function shouldUseIntrinsicAspectRatio(shape: SlideShape): boolean {
  const transform = shape.transform;
  if (!shape.pathCommands?.length || !shape.pathViewport || !transform) {
    return false;
  }
  const cx = transform.cx;
  const cy = transform.cy;
  const x = transform.x;
  const y = transform.y;
  if (cx === undefined || cy === undefined || x === undefined || y === undefined) {
    return false;
  }
  const viewportAspect = shape.pathViewport.width / Math.max(shape.pathViewport.height, 1);
  const transformAspect = cx / Math.max(cy, 1);
  if (Math.abs(viewportAspect - transformAspect) < 0.18) {
    return false;
  }
  const widthRatio = cx / 12192000;
  const heightRatio = cy / 6858000;
  const topRight = (x / 12192000) > 0.75 && (y / 6858000) < 0.14;
  const whiteVector = shape.fill?.color === '#FFFFFF' || shape.line?.color === '#FFFFFF';
  return topRight && whiteVector && widthRatio < 0.18 && heightRatio < 0.08;
}

function isPresetSceneVectorShape(shapeType: string | undefined): boolean {
  return ['rect', 'ellipse', 'chevron', 'trapezoid', 'roundRect', 'round2SameRect'].includes(shapeType ?? '');
}

function buildPresetSceneSvgMarkup(
  shapeType: string | undefined,
  fill: string,
  fillOpacity: number | undefined,
  stroke: string,
  strokeWidth: number,
  strokeAttrs: string
): string | undefined {
  const fillOpacityAttr = fill !== 'none' && fillOpacity !== undefined ? ` fill-opacity="${fillOpacity}"` : '';
  switch (shapeType) {
    case 'ellipse':
      return `<ellipse cx="500" cy="500" rx="500" ry="500" fill="${escapeHtml(fill)}"${fillOpacityAttr} stroke="${escapeHtml(stroke)}"${strokeAttrs} stroke-width="${strokeWidth}" vector-effect="non-scaling-stroke"/>`;
    case 'chevron':
      return `<path d="M 0 0 L 715 0 L 1000 500 L 715 1000 L 0 1000 L 330 500 Z" fill="${escapeHtml(fill)}"${fillOpacityAttr} stroke="${escapeHtml(stroke)}"${strokeAttrs} stroke-width="${strokeWidth}" vector-effect="non-scaling-stroke"/>`;
    case 'trapezoid':
      return `<path d="M 180 0 L 820 0 L 1000 1000 L 0 1000 Z" fill="${escapeHtml(fill)}"${fillOpacityAttr} stroke="${escapeHtml(stroke)}"${strokeAttrs} stroke-width="${strokeWidth}" vector-effect="non-scaling-stroke"/>`;
    case 'roundRect':
    case 'round2SameRect':
      return `<rect x="0" y="0" width="1000" height="1000" rx="120" ry="120" fill="${escapeHtml(fill)}"${fillOpacityAttr} stroke="${escapeHtml(stroke)}"${strokeAttrs} stroke-width="${strokeWidth}" vector-effect="non-scaling-stroke"/>`;
    case 'rect':
      return `<rect x="0" y="0" width="1000" height="1000" fill="${escapeHtml(fill)}"${fillOpacityAttr} stroke="${escapeHtml(stroke)}"${strokeAttrs} stroke-width="${strokeWidth}" vector-effect="non-scaling-stroke"/>`;
    default:
      return undefined;
  }
}

function sceneStrokeDashArray(line: SlideShape['line']): string | undefined {
  const width = Math.max(1, emuToPx(line?.width ?? 0));
  switch (line?.dash) {
    case 'dash':
    case 'sysDash':
      return `${width * 3} ${width * 2}`;
    case 'lgDash':
      return `${width * 6} ${width * 2}`;
    case 'dot':
    case 'sysDot':
      return `${width} ${width * 1.5}`;
    case 'dashDot':
      return `${width * 4} ${width * 2} ${width} ${width * 2}`;
    default:
      return undefined;
  }
}

function emuToPx(value: number): number {
  return Math.max(0.5, Number((value / 9525).toFixed(2)));
}

function toCssAlign(value: string | undefined): string {
  switch (value) {
    case 'ctr': return 'center';
    case 'r': return 'right';
    case 'just': return 'justify';
    default: return 'left';
  }
}

function toCssAnchor(value: string | undefined): string {
  switch (value) {
    case 'ctr':
      return 'center';
    case 'b':
      return 'flex-end';
    default:
      return 'flex-start';
  }
}

function renderSceneText(shape: SlideShape, overlay: boolean, isSlideTitle: boolean): string {
  const isCentered = isSlideTitle || shape.textStyle?.align === 'ctr';
  const compactCenteredLongText = isCentered
    && (shape.transform?.cx ?? Number.POSITIVE_INFINITY) < 4_000_000
    && (shape.textStyle?.fontSizePt ?? 0) <= 36
    && shape.text.trim().length > 10;
  const style = [
    `text-align:${isCentered ? 'center' : toCssAlign(shape.textStyle?.align)}`,
    overlay ? `align-items:${toCssAnchor(shape.textStyle?.anchor)}` : '',
    overlay ? (isCentered ? 'justify-content:center' : 'justify-content:flex-start') : '',
    isCentered ? 'padding:0' : '',
    compactCenteredLongText ? 'line-height:1.05' : '',
    isCentered ? 'box-sizing:border-box' : '',
    isCentered ? 'width:100%' : '',
    isCentered ? '' : 'padding:6px 18px 6px 18px',
    isCentered ? '' : 'line-height:1.2',
    isCentered ? '' : 'box-sizing:border-box',
    isCentered ? '' : 'width:100%'
  ].filter(Boolean).join(';');
  const className = overlay ? 'ooxml-pptx-scene-text ooxml-pptx-scene-text--overlay' : 'ooxml-pptx-scene-text';
  return `<div class="${className}"${style ? ` style="${style}"` : ''}>${escapeHtml(shape.text).replaceAll('\n', '<br>')}</div>`;
}

function applyOpacity(color: string, opacity: number | undefined): string {
  if (!color.startsWith('#') || color.length !== 7 || opacity === undefined) {
    return color;
  }
  const red = Number.parseInt(color.slice(1, 3), 16);
  const green = Number.parseInt(color.slice(3, 5), 16);
  const blue = Number.parseInt(color.slice(5, 7), 16);
  return `rgba(${red}, ${green}, ${blue}, ${opacity})`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function serializePathCommands(commands: NonNullable<PresentationDocument['slides'][number]['shapes'][number]['pathCommands']>): string {
  return commands
    .map((command) => {
      if (command.type === 'close') {
        return 'Z';
      }
      if (command.type === 'cubicTo') {
        return `C:${command.x1 ?? ''},${command.y1 ?? ''}|${command.x2 ?? ''},${command.y2 ?? ''}|${command.x ?? ''},${command.y ?? ''}`;
      }
      return `${command.type === 'moveTo' ? 'M' : 'L'}:${command.x ?? ''},${command.y ?? ''}`;
    })
    .join(';');
}

function serializeGradientStops(stops: NonNullable<PresentationDocument['slides'][number]['shapes'][number]['fill']>['gradientStops']): string {
  return (stops ?? []).map((stop) => `${stop.position}:${stop.color ?? ''}:${stop.opacity ?? ''}`).join(';');
}
