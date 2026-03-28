import { relationshipsFor, updatePackagePartText, upsertRelationship } from '@ooxml/core';
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

export function setDocxRevisionMetadata(
  editor: OfficeEditor<DocxDocument>,
  storyKind: 'document' | 'header' | 'footer',
  storyOccurrence: number,
  paragraphIndex: number,
  revisionIndex: number,
  patch: { id?: string; kind?: 'insertion' | 'deletion'; author?: string; date?: string; text?: string }
): DocxDocument {
  return editor.transaction((draft) => {
    const story = draft.stories.filter((entry) => entry.kind === storyKind)[storyOccurrence];
    const paragraph = story?.paragraphs[paragraphIndex];
    const revision = paragraph?.revisions[revisionIndex];
    if (!paragraph || !revision) {
      return;
    }

    if (patch.author !== undefined) {
      revision.author = patch.author;
    }
    if (patch.date !== undefined) {
      revision.date = patch.date;
    }
    if (patch.id !== undefined) {
      revision.id = patch.id;
    }
    if (patch.kind !== undefined) {
      revision.kind = patch.kind;
    }
    if (patch.text !== undefined) {
      revision.text = patch.text;
    }

    paragraph.text = [
      ...paragraph.runs.map((run) => run.text),
      ...paragraph.revisions.filter((entry) => entry.kind === 'insertion').map((entry) => entry.text)
    ].join('');
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

export function addDocxComment(editor: OfficeEditor<DocxDocument>, comment: { id: string; author?: string; text: string }): DocxDocument {
  return editor.transaction((draft) => {
    const commentsUri = ensureDocxCommentsPart(draft);
    if (!commentsUri) {
      return;
    }

    if (!draft.comments.find((entry) => entry.id === comment.id)) {
      draft.comments.push({
        id: comment.id,
        author: comment.author,
        text: comment.text
      });
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

function ensureDocxCommentsPart(document: DocxDocument): string | undefined {
  const mainDocumentUri = document.packageGraph.rootDocumentUri ?? '/word/document.xml';
  const existingRelationship = relationshipsFor(document.packageGraph, mainDocumentUri).find((relationship) => relationship.type.includes('/comments'));
  if (existingRelationship?.resolvedTarget && document.packageGraph.parts[existingRelationship.resolvedTarget]) {
    return existingRelationship.resolvedTarget;
  }

  const commentsUri = '/word/comments.xml';
  updatePackagePartText(
    document.packageGraph,
    commentsUri,
    `<?xml version="1.0" encoding="UTF-8"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>`,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
  );
  upsertRelationship(document.packageGraph, mainDocumentUri, {
    id: existingRelationship?.id ?? nextRelationshipId(relationshipsFor(document.packageGraph, mainDocumentUri)),
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    target: 'comments.xml',
    targetMode: 'Internal'
  });
  return commentsUri;
}

function nextRelationshipId(relationships: ReturnType<typeof relationshipsFor>): string {
  let candidateIndex = relationships.length + 1;
  let candidate = `rId${candidateIndex}`;
  const existingIds = new Set(relationships.map((relationship) => relationship.id));
  while (existingIds.has(candidate)) {
    candidateIndex += 1;
    candidate = `rId${candidateIndex}`;
  }
  return candidate;
}

export function setDocxSectionReferenceType(
  editor: OfficeEditor<DocxDocument>,
  sectionIndex: number,
  referenceKind: 'header' | 'footer',
  referenceIndex: number,
  type: 'default' | 'first' | 'even'
): DocxDocument {
  return editor.transaction((draft) => {
    const section = draft.sections[sectionIndex];
    const reference = (referenceKind === 'header' ? section?.headerReferences : section?.footerReferences)?.[referenceIndex];
    if (reference) {
      reference.type = type;
    }
  });
}

export function setDocxSectionReferenceTarget(
  editor: OfficeEditor<DocxDocument>,
  sectionIndex: number,
  referenceKind: 'header' | 'footer',
  referenceIndex: number,
  targetUri: string
): DocxDocument {
  return editor.transaction((draft) => {
    const section = draft.sections[sectionIndex];
    const reference = (referenceKind === 'header' ? section?.headerReferences : section?.footerReferences)?.[referenceIndex];
    if (reference) {
      reference.targetUri = targetUri;
    }
  });
}
