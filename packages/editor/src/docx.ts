import type { DocxDocument } from '@ooxml/docx';

import type { OfficeEditor } from './types';

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


export function setDocxCommentText(editor: OfficeEditor<DocxDocument>, commentId: string, text: string): DocxDocument {
  return editor.transaction((draft) => {
    const comment = draft.comments.find((entry) => entry.id === commentId);
    if (comment) {
      comment.text = text;
    }
  });
}
