import type { DocxDocument } from '@ooxml/docx';
import type { PresentationDocument } from '@ooxml/pptx';
import type { XlsxWorkbook } from '@ooxml/xlsx';

export type RenderableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export interface RenderOptions {
  activeSheetIndex?: number;
  activeSlideIndex?: number;
  showComments?: boolean;
}
