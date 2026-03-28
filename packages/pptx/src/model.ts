import type { PackageGraph } from '@ooxml/core';

export interface SlideShape {
  id: string;
  name?: string;
  text: string;
  placeholderType?: string;
}

export interface PresentationTheme {
  uri: string;
  name?: string;
  colorSchemeName?: string;
  majorLatinFont?: string;
  minorLatinFont?: string;
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
}

export interface PresentationDocument {
  kind: 'pptx';
  packageGraph: PackageGraph;
  slides: PresentationSlide[];
  size: { cx: number; cy: number };
  themes: Record<string, PresentationTheme>;
}
