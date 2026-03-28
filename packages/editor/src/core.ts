import { serializeOfficeDocument } from '@ooxml/serializer';

import type { EditableOfficeDocument, OfficeEditor } from './types';

export function createOfficeEditor<TDocument extends EditableOfficeDocument>(document: TDocument): OfficeEditor<TDocument> {
  let current = structuredClone(document) as TDocument;
  const undoStack: TDocument[] = [];
  const redoStack: TDocument[] = [];

  return {
    get document() {
      return current;
    },
    transaction(mutator) {
      const draft = structuredClone(current) as TDocument;
      mutator(draft);
      undoStack.push(structuredClone(current) as TDocument);
      current = draft;
      redoStack.length = 0;
      return current;
    },
    undo() {
      if (undoStack.length === 0) {
        return current;
      }

      redoStack.push(structuredClone(current) as TDocument);
      current = undoStack.pop() as TDocument;
      return current;
    },
    redo() {
      if (redoStack.length === 0) {
        return current;
      }

      undoStack.push(structuredClone(current) as TDocument);
      current = redoStack.pop() as TDocument;
      return current;
    },
    serialize() {
      return serializeOfficeDocument(current);
    }
  };
}
