import type { PackageGraph } from '@ooxml/core';

export interface SlideShape {
  id: string;
  name?: string;
  text: string;
}

export interface PresentationSlide {
  title: string;
  uri: string;
  notesUri?: string;
  notesText: string;
  shapes: SlideShape[];
}

export interface PresentationDocument {
  kind: 'pptx';
  packageGraph: PackageGraph;
  slides: PresentationSlide[];
  size: { cx: number; cy: number };
}
