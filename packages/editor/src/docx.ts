import type { DocxDocument } from '@ooxml/docx';

import type { OfficeEditor } from './types';

export function replaceDocxParagraphText(editor: OfficeEditor<DocxDocument>, storyIndex: number, paragraphIndex: number, text: string): DocxDocument {
  return replaceDocxStoryParagraphText(editor, editor.document.stories[storyIndex]?.kind ?? 'document', 0, paragraphIndex, text, storyIndex);
}

export function replaceDocxStoryParagraphText(
  editor: OfficeEditor<DocxDocument>,
  storyKind: 'document' | 'header' | 'footer',
  storyOccurrence: number,
  paragraphIndex: number,
  text: string,
  storyIndexHint?: number
): DocxDocument {
  return editor.transaction((draft) => {
    const story =
      typeof storyIndexHint === 'number' && draft.stories[storyIndexHint]?.kind === storyKind
        ? draft.stories[storyIndexHint]
        : draft.stories.filter((entry) => entry.kind === storyKind)[storyOccurrence];
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
