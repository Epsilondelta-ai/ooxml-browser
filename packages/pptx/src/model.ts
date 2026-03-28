import type { PackageGraph } from '@ooxml/core';

export interface SlideShapeMedia {
  type: 'image';
  targetUri?: string;
}

export interface SlideShape {
  id: string;
  name?: string;
  text: string;
  placeholderType?: string;
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
