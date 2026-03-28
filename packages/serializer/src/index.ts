import { serializeDocx } from './docx';
import { serializePptx } from './pptx';
import { serializeXlsx } from './xlsx';

import type { SerializableOfficeDocument } from './types';

export type { SerializableOfficeDocument } from './types';

export function serializeOfficeDocument(document: SerializableOfficeDocument): Uint8Array {
  switch (document.kind) {
    case 'docx':
      return serializeDocx(document);
    case 'xlsx':
      return serializeXlsx(document);
    case 'pptx':
      return serializePptx(document);
  }
}
