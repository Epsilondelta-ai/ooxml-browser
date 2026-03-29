import type { PackageGraph } from '@ooxml/core';

export interface SlideShapeTransform {
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface PresentationFill {
  kind: 'solid' | 'none' | 'image' | 'gradient';
  color?: string;
  opacity?: number;
  targetUri?: string;
  gradientStops?: { position: number; color?: string; opacity?: number }[];
  angleDeg?: number;
}

export interface PresentationLine {
  kind: 'solid' | 'none';
  color?: string;
  opacity?: number;
  width?: number;
}

export interface PresentationTextStyle {
  color?: string;
  fontSizePt?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  align?: string;
}

export interface PresentationPathCommand {
  type: 'moveTo' | 'lineTo' | 'cubicTo' | 'close';
  x?: number;
  y?: number;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
}

export interface SlideShapeMedia {
  type: 'image' | 'embeddedObject';
  targetUri?: string;
  progId?: string;
}

export interface SlideShape {
  id: string;
  name?: string;
  text: string;
  placeholderType?: string;
  placeholderIndex?: string;
  shapeType?: string;
  transform?: SlideShapeTransform;
  fill?: PresentationFill;
  line?: PresentationLine;
  textStyle?: PresentationTextStyle;
  pathCommands?: PresentationPathCommand[];
  pathViewport?: {
    width: number;
    height: number;
  };
  media?: SlideShapeMedia;
}

export interface PresentationTheme {
  uri: string;
  name?: string;
  colorSchemeName?: string;
  majorLatinFont?: string;
  minorLatinFont?: string;
  colors?: Record<string, string>;
}

export interface PresentationComment {
  author?: string;
  text: string;
  index: number;
}

export interface PresentationTransition {
  type?: string;
  speed?: string;
  advanceOnClick?: boolean;
  advanceAfterMs?: number;
}

export interface PresentationTimingNode {
  nodeType: string;
  concurrent?: boolean;
  nextAction?: string;
  previousAction?: string;
  presetClass?: string;
  presetId?: string;
  id?: string;
  duration?: string;
  repeatDuration?: string;
  repeatCount?: string;
  restart?: string;
  fill?: string;
  autoReverse?: boolean;
  acceleration?: string;
  deceleration?: string;
  triggerEvent?: string;
  triggerDelay?: string;
  triggerShapeId?: string;
  endTriggerEvent?: string;
  endTriggerDelay?: string;
  endTriggerShapeId?: string;
  targetShapeId?: string;
  colorSpace?: string;
  colorDirection?: string;
  motionOrigin?: string;
  motionPath?: string;
  motionPathEditMode?: string;
  commandName?: string;
  commandType?: string;
}

export interface PresentationTiming {
  nodeCount: number;
  nodes: PresentationTimingNode[];
}

export interface PresentationSlide {
  title: string;
  uri: string;
  notesUri?: string;
  notesText: string;
  layoutUri?: string;
  layoutName?: string;
  masterUri?: string;
  masterName?: string;
  themeUri?: string;
  background?: PresentationFill;
  transition?: PresentationTransition;
  timing?: PresentationTiming;
  shapes: SlideShape[];
  comments: PresentationComment[];
}

export interface PresentationDocument {
  kind: 'pptx';
  packageGraph: PackageGraph;
  slides: PresentationSlide[];
  size: { cx: number; cy: number };
  themes: Record<string, PresentationTheme>;
}
