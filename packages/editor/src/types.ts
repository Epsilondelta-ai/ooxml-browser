import type { DocxDocument } from '@ooxml/docx';
import type { PresentationDocument } from '@ooxml/pptx';
import type { XlsxWorkbook } from '@ooxml/xlsx';

export type EditableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export interface OfficeEditor<TDocument extends EditableOfficeDocument> {
  document: TDocument;
  transaction(mutator: (draft: TDocument) => void): TDocument;
  undo(): TDocument;
  redo(): TDocument;
  serialize(): Uint8Array;
}
