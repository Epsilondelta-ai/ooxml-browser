import type { PresentationDocument } from '@ooxml/pptx';

import type { RenderOptions } from './types';

export function renderPresentation(presentation: PresentationDocument, options: RenderOptions): string {
  const slideIndex = Math.min(options.activeSlideIndex ?? 0, Math.max(presentation.slides.length - 1, 0));
  const slide = presentation.slides[slideIndex];

  if (!slide) {
    return '<section class="ooxml-render ooxml-render--pptx">No slides available</section>';
  }

  const theme = slide.themeUri ? presentation.themes[slide.themeUri] : undefined;
  const shapeMarkup = slide.shapes.map((shape) => `<div class="ooxml-pptx-shape"${shape.placeholderType ? ` data-placeholder-type="${escapeHtml(shape.placeholderType)}"` : ''}${shape.placeholderIndex ? ` data-placeholder-index="${escapeHtml(shape.placeholderIndex)}"` : ''}${shape.shapeType ? ` data-shape-type="${escapeHtml(shape.shapeType)}"` : ''}${shape.fill?.color ? ` data-fill-color="${escapeHtml(shape.fill.color)}"` : ''}${shape.fill?.opacity !== undefined ? ` data-fill-opacity="${shape.fill.opacity}"` : ''}${shape.fill?.targetUri ? ` data-fill-image-uri="${escapeHtml(shape.fill.targetUri)}"` : ''}${shape.fill?.gradientStops?.length ? ` data-fill-gradient-stops="${escapeHtml(serializeGradientStops(shape.fill.gradientStops))}"` : ''}${shape.fill?.angleDeg !== undefined ? ` data-fill-gradient-angle="${shape.fill.angleDeg}"` : ''}${shape.line?.color ? ` data-line-color="${escapeHtml(shape.line.color)}"` : ''}${shape.line?.opacity !== undefined ? ` data-line-opacity="${shape.line.opacity}"` : ''}${shape.line?.width !== undefined ? ` data-line-width="${shape.line.width}"` : ''}${shape.textStyle?.color ? ` data-text-color="${escapeHtml(shape.textStyle.color)}"` : ''}${shape.textStyle?.fontSizePt !== undefined ? ` data-font-size-pt="${shape.textStyle.fontSizePt}"` : ''}${shape.textStyle?.fontFamily ? ` data-font-family="${escapeHtml(shape.textStyle.fontFamily)}"` : ''}${shape.textStyle?.bold !== undefined ? ` data-font-bold="${shape.textStyle.bold}"` : ''}${shape.textStyle?.italic !== undefined ? ` data-font-italic="${shape.textStyle.italic}"` : ''}${shape.textStyle?.align ? ` data-text-align="${escapeHtml(shape.textStyle.align)}"` : ''}${shape.pathCommands ? ` data-path-commands="${escapeHtml(serializePathCommands(shape.pathCommands))}"` : ''}${shape.media?.targetUri ? ` data-media-uri="${escapeHtml(shape.media.targetUri)}"` : ''}${shape.media?.type ? ` data-media-type="${escapeHtml(shape.media.type)}"` : ''}${shape.media?.progId ? ` data-prog-id="${escapeHtml(shape.media.progId)}"` : ''}${shape.transform?.x !== undefined ? ` data-x="${shape.transform.x}"` : ''}${shape.transform?.y !== undefined ? ` data-y="${shape.transform.y}"` : ''}${shape.transform?.cx !== undefined ? ` data-cx="${shape.transform.cx}"` : ''}${shape.transform?.cy !== undefined ? ` data-cy="${shape.transform.cy}"` : ''}${shape.transform?.rotationDeg !== undefined ? ` data-rotation-deg="${shape.transform.rotationDeg}"` : ''}${shape.transform?.flipH ? ' data-flip-h="true"' : ''}${shape.transform?.flipV ? ' data-flip-v="true"' : ''}><h3>${escapeHtml(shape.name ?? 'Shape')}</h3><p>${escapeHtml(shape.text || (shape.media?.type === 'embeddedObject' ? '[embedded-object]' : shape.media ? '[image]' : ''))}</p></div>`).join('');
  const notes = slide.notesText ? `<aside class="ooxml-pptx-notes">${escapeHtml(slide.notesText)}</aside>` : '';
  const commentsMarkup = slide.comments.length ? `<ul class="ooxml-pptx-comments">${slide.comments.map((comment) => `<li data-comment-index="${comment.index}">${escapeHtml(comment.text)}${comment.author ? ` — ${escapeHtml(comment.author)}` : ''}</li>`).join('')}</ul>` : '';
  const timingMarkup = slide.transition || slide.timing ? `<dl class="ooxml-pptx-timing">${slide.transition?.type ? `<dt>Transition</dt><dd data-transition-type="${escapeHtml(slide.transition.type)}">${escapeHtml(slide.transition.type)}${slide.transition.speed ? ` (${escapeHtml(slide.transition.speed)})` : ''}${slide.transition.advanceOnClick !== undefined ? ` click:${escapeHtml(String(slide.transition.advanceOnClick))}` : ''}${slide.transition.advanceAfterMs !== undefined ? ` after:${escapeHtml(String(slide.transition.advanceAfterMs))}` : ''}</dd>` : ''}${slide.timing ? `<dt>Timing nodes</dt><dd data-timing-count="${slide.timing.nodeCount}">${slide.timing.nodes.map((node) => `${escapeHtml(node.nodeType)}${node.presetClass ? `:${escapeHtml(node.presetClass)}` : ''}${node.concurrent !== undefined ? ` concurrent:${escapeHtml(String(node.concurrent))}` : ''}${node.nextAction ? ` next:${escapeHtml(node.nextAction)}` : ''}${node.previousAction ? ` prev:${escapeHtml(node.previousAction)}` : ''}${node.id ? `#${escapeHtml(node.id)}` : ''}${node.duration ? `@${escapeHtml(node.duration)}` : ''}${node.repeatDuration ? ` repeatDur:${escapeHtml(node.repeatDuration)}` : ''}${node.repeatCount ? `×${escapeHtml(node.repeatCount)}` : ''}${node.restart ? ` restart:${escapeHtml(node.restart)}` : ''}${node.fill ? ` fill:${escapeHtml(node.fill)}` : ''}${node.autoReverse !== undefined ? ` autoRev:${escapeHtml(String(node.autoReverse))}` : ''}${node.acceleration ? ` accel:${escapeHtml(node.acceleration)}` : ''}${node.deceleration ? ` decel:${escapeHtml(node.deceleration)}` : ''}${node.colorSpace ? ` clr:${escapeHtml(node.colorSpace)}` : ''}${node.colorDirection ? ` dir:${escapeHtml(node.colorDirection)}` : ''}${node.motionOrigin ? ` origin:${escapeHtml(node.motionOrigin)}` : ''}${node.motionPath ? ` path:${escapeHtml(node.motionPath)}` : ''}${node.motionPathEditMode ? ` pathEdit:${escapeHtml(node.motionPathEditMode)}` : ''}${node.commandName ? ` cmd:${escapeHtml(node.commandName)}` : ''}${node.commandType ? ` cmdType:${escapeHtml(node.commandType)}` : ''}${node.triggerEvent ? `!${escapeHtml(node.triggerEvent)}` : ''}${node.triggerDelay ? `+${escapeHtml(node.triggerDelay)}` : ''}${node.triggerShapeId ? `^${escapeHtml(node.triggerShapeId)}` : ''}${node.endTriggerEvent ? ` ~${escapeHtml(node.endTriggerEvent)}` : ''}${node.endTriggerDelay ? `=${escapeHtml(node.endTriggerDelay)}` : ''}${node.endTriggerShapeId ? `^${escapeHtml(node.endTriggerShapeId)}` : ''}${node.targetShapeId ? `->${escapeHtml(node.targetShapeId)}` : ''}`).join(', ')}</dd>` : ''}</dl>` : '';
  const inheritanceMarkup = `<dl class="ooxml-pptx-inheritance"><dt>Layout</dt><dd>${escapeHtml(slide.layoutName ?? slide.layoutUri ?? 'none')}</dd><dt>Master</dt><dd>${escapeHtml(slide.masterName ?? slide.masterUri ?? 'none')}</dd><dt>Theme</dt><dd>${escapeHtml(theme?.name ?? theme?.colorSchemeName ?? slide.themeUri ?? 'none')}</dd>${theme?.majorLatinFont ? `<dt>Major font</dt><dd>${escapeHtml(theme.majorLatinFont)}</dd>` : ''}${theme?.minorLatinFont ? `<dt>Minor font</dt><dd>${escapeHtml(theme.minorLatinFont)}</dd>` : ''}</dl>`;

  return `<section class="ooxml-render ooxml-render--pptx" data-presentation-cx="${presentation.size.cx}" data-presentation-cy="${presentation.size.cy}"${slide.background?.color ? ` data-background-color="${escapeHtml(slide.background.color)}"` : ''}${slide.background?.opacity !== undefined ? ` data-background-opacity="${slide.background.opacity}"` : ''}${slide.background?.targetUri ? ` data-background-image-uri="${escapeHtml(slide.background.targetUri)}"` : ''}${slide.layoutUri ? ` data-layout-uri="${escapeHtml(slide.layoutUri)}"` : ''}${slide.masterUri ? ` data-master-uri="${escapeHtml(slide.masterUri)}"` : ''}${slide.themeUri ? ` data-theme-uri="${escapeHtml(slide.themeUri)}"` : ''}><header><h2>${escapeHtml(slide.title)}</h2><p>${presentation.size.cx} × ${presentation.size.cy}</p></header>${inheritanceMarkup}${timingMarkup}${shapeMarkup}${commentsMarkup}${notes}</section>`;
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
