import type { DocxDocument } from '@ooxml/docx';
import type { PresentationDocument } from '@ooxml/pptx';
import type { XlsxWorkbook } from '@ooxml/xlsx';

export type SerializableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;
