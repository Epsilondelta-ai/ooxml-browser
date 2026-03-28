import type { PresentationDocument } from '@ooxml/pptx';

import type { OfficeEditor } from './types';

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
