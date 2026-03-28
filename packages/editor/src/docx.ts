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

export function setDocxParagraphStyle(
  editor: OfficeEditor<DocxDocument>,
  storyKind: 'document' | 'header' | 'footer',
  storyOccurrence: number,
  paragraphIndex: number,
  styleId: string | undefined
): DocxDocument {
  return editor.transaction((draft) => {
    const story = draft.stories.filter((entry) => entry.kind === storyKind)[storyOccurrence];
    const paragraph = story?.paragraphs[paragraphIndex];
    if (paragraph) {
      paragraph.styleId = styleId;
    }
  });
}

export function setDocxParagraphNumbering(
  editor: OfficeEditor<DocxDocument>,
  storyKind: 'document' | 'header' | 'footer',
  storyOccurrence: number,
  paragraphIndex: number,
  numbering: { numId: string; level: number } | undefined
): DocxDocument {
  return editor.transaction((draft) => {
    const story = draft.stories.filter((entry) => entry.kind === storyKind)[storyOccurrence];
    const paragraph = story?.paragraphs[paragraphIndex];
    if (paragraph) {
      paragraph.numbering = numbering ? { ...numbering } : undefined;
    }
  });
}

export function setDocxParagraphRunStyle(
  editor: OfficeEditor<DocxDocument>,
  storyKind: 'document' | 'header' | 'footer',
  storyOccurrence: number,
  paragraphIndex: number,
  runIndex: number,
  style: { bold?: boolean; italic?: boolean }
): DocxDocument {
  return editor.transaction((draft) => {
    const story = draft.stories.filter((entry) => entry.kind === storyKind)[storyOccurrence];
    const paragraph = story?.paragraphs[paragraphIndex];
    const run = paragraph?.runs[runIndex];
    if (run) {
      run.bold = style.bold ?? run.bold;
      run.italic = style.italic ?? run.italic;
    }
  });
}

export function setDocxTableCellText(
  editor: OfficeEditor<DocxDocument>,
  storyKind: 'document' | 'header' | 'footer',
  storyOccurrence: number,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  text: string
): DocxDocument {
  return editor.transaction((draft) => {
    const story = draft.stories.filter((entry) => entry.kind === storyKind)[storyOccurrence];
    const cell = story?.tables[tableIndex]?.rows[rowIndex]?.cells[cellIndex];
    if (cell) {
      cell.text = text;
    }
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

export function setDocxCommentAuthor(editor: OfficeEditor<DocxDocument>, commentId: string, author: string): DocxDocument {
  return editor.transaction((draft) => {
    const comment = draft.comments.find((entry) => entry.id === commentId);
    if (comment) {
      comment.author = author;
    }
  });
}

export function setDocxSectionLayout(editor: OfficeEditor<DocxDocument>, sectionIndex: number, layout: { pageSize?: { width: number; height: number }; pageMargins?: { top: number; right: number; bottom: number; left: number } }): DocxDocument {
  return editor.transaction((draft) => {
    const section = draft.sections[sectionIndex];
    if (!section) {
      return;
    }

    if (layout.pageSize) {
      section.pageSize = { ...layout.pageSize };
    }
    if (layout.pageMargins) {
      section.pageMargins = { ...layout.pageMargins };
    }
  });
}
