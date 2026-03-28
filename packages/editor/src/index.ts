import type { DocxDocument } from '@ooxml/docx';
import type { PresentationDocument } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';
import type { XlsxWorkbook } from '@ooxml/xlsx';

export type EditableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export interface OfficeEditor<TDocument extends EditableOfficeDocument> {
  document: TDocument;
  transaction(mutator: (draft: TDocument) => void): TDocument;
  undo(): TDocument;
  redo(): TDocument;
  serialize(): Uint8Array;
}

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

export function replaceDocxParagraphText(editor: OfficeEditor<DocxDocument>, storyIndex: number, paragraphIndex: number, text: string): DocxDocument {
  return editor.transaction((draft) => {
    const story = draft.stories[storyIndex];
    const paragraph = story?.paragraphs[paragraphIndex];
    if (!paragraph) {
      return;
    }

    paragraph.text = text;
    if (paragraph.runs.length === 0) {
      paragraph.runs.push({ text, bold: false, italic: false });
      return;
    }

    paragraph.runs = [{ ...paragraph.runs[0], text }];
  });
}

export function setWorkbookCellValue(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, value: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    for (const row of sheet.rows) {
      const cell = row.cells.find((entry) => entry.reference === reference);
      if (!cell) {
        continue;
      }

      cell.value = value;
      cell.type = Number.isNaN(Number(value)) ? 's' : 'n';
      cell.formula = undefined;
      return;
    }
  });
}

export function setPresentationShapeText(editor: OfficeEditor<PresentationDocument>, slideIndex: number, shapeIndex: number, text: string): PresentationDocument {
  return editor.transaction((draft) => {
    const shape = draft.slides[slideIndex]?.shapes[shapeIndex];
    if (shape) {
      shape.text = text;
    }
  });
}

export function setPresentationNotesText(editor: OfficeEditor<PresentationDocument>, slideIndex: number, text: string): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (!slide) {
      return;
    }

    const notesUri = slide.notesUri;
    if (!notesUri || !draft.packageGraph.parts[notesUri]) {
      return;
    }

    slide.notesText = text;
  });
}
