import type { PackageGraph } from '@ooxml/core';

export interface SlideShapeTransform {
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
}

export interface SlideShapeMedia {
  type: 'image';
  targetUri?: string;
}

export interface SlideShape {
  id: string;
  name?: string;
  text: string;
  placeholderType?: string;
  transform?: SlideShapeTransform;
  media?: SlideShapeMedia;
}

export interface PresentationTheme {
  uri: string;
  name?: string;
  colorSchemeName?: string;
  majorLatinFont?: string;
  minorLatinFont?: string;
}

export interface PresentationComment {
  author?: string;
  text: string;
  index: number;
}

export interface PresentationTransition {
  type?: string;
  speed?: string;
}

export interface PresentationTimingNode {
  nodeType: string;
  presetClass?: string;
  presetId?: string;
  id?: string;
  duration?: string;
  repeatCount?: string;
  restart?: string;
  fill?: string;
  acceleration?: string;
  deceleration?: string;
  triggerEvent?: string;
  triggerDelay?: string;
  endTriggerEvent?: string;
  endTriggerDelay?: string;
  targetShapeId?: string;
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
