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


export function setPresentationCommentText(editor: OfficeEditor<PresentationDocument>, slideIndex: number, commentIndex: number, text: string): PresentationDocument {
  return editor.transaction((draft) => {
    const comment = draft.slides[slideIndex]?.comments[commentIndex];
    if (comment) {
      comment.text = text;
    }
  });
}

export function setPresentationShapeTransform(editor: OfficeEditor<PresentationDocument>, slideIndex: number, shapeIndex: number, transform: { x?: number; y?: number; cx?: number; cy?: number }): PresentationDocument {
  return editor.transaction((draft) => {
    const shape = draft.slides[slideIndex]?.shapes[shapeIndex];
    if (shape) {
      shape.transform = { ...shape.transform, ...transform };
    }
  });
}

export function setPresentationTransition(editor: OfficeEditor<PresentationDocument>, slideIndex: number, transition: { type?: string; speed?: string } | undefined): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (slide) {
      slide.transition = transition;
    }
  });
}

export function setPresentationTimingNodes(editor: OfficeEditor<PresentationDocument>, slideIndex: number, nodes: Array<{ nodeType: string; presetClass?: string; presetId?: string }>): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (slide) {
      slide.timing = { nodeCount: nodes.length, nodes: [...nodes] };
    }
  });
}
