import { relationshipsFor, updatePackagePartText, upsertRelationship } from '@ooxml/core';
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

    const notesUri = ensurePresentationNotesPart(draft, slideIndex);
    if (!notesUri) {
      return;
    }

    slide.notesUri = notesUri;
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

export function addPresentationComment(editor: OfficeEditor<PresentationDocument>, slideIndex: number, comment: { author?: string; text: string }): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (!slide) {
      return;
    }

    const commentsUri = ensurePresentationCommentsPart(draft, slideIndex);
    if (!commentsUri) {
      return;
    }

    slide.comments.push({
      author: comment.author,
      text: comment.text,
      index: slide.comments.length
    });
  });
}

export function setPresentationCommentAuthor(editor: OfficeEditor<PresentationDocument>, slideIndex: number, commentIndex: number, author: string): PresentationDocument {
  return editor.transaction((draft) => {
    const comment = draft.slides[slideIndex]?.comments[commentIndex];
    if (comment) {
      comment.author = author;
    }
  });
}

export function removePresentationComment(editor: OfficeEditor<PresentationDocument>, slideIndex: number, commentIndex: number): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (!slide) {
      return;
    }

    slide.comments = slide.comments
      .filter((_comment, index) => index !== commentIndex)
      .map((comment, index) => ({ ...comment, index }));
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

export function setPresentationShapePlaceholderType(editor: OfficeEditor<PresentationDocument>, slideIndex: number, shapeIndex: number, placeholderType: string | undefined): PresentationDocument {
  return editor.transaction((draft) => {
    const shape = draft.slides[slideIndex]?.shapes[shapeIndex];
    if (shape) {
      shape.placeholderType = placeholderType;
    }
  });
}

export function setPresentationShapeName(editor: OfficeEditor<PresentationDocument>, slideIndex: number, shapeIndex: number, name: string): PresentationDocument {
  return editor.transaction((draft) => {
    const shape = draft.slides[slideIndex]?.shapes[shapeIndex];
    if (shape) {
      shape.name = name;
    }
  });
}

export function setPresentationImageTarget(editor: OfficeEditor<PresentationDocument>, slideIndex: number, shapeIndex: number, targetUri: string): PresentationDocument {
  return editor.transaction((draft) => {
    const shape = draft.slides[slideIndex]?.shapes[shapeIndex];
    if (shape?.media?.type === 'image') {
      shape.media = {
        ...shape.media,
        targetUri
      };
    }
  });
}

export function setPresentationSlideLayout(editor: OfficeEditor<PresentationDocument>, slideIndex: number, layoutUri: string): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (slide) {
      slide.layoutUri = layoutUri;
      slide.layoutName = undefined;
      slide.masterUri = undefined;
      slide.masterName = undefined;
      slide.themeUri = undefined;
    }
  });
}

export function setPresentationSlideMaster(editor: OfficeEditor<PresentationDocument>, slideIndex: number, masterUri: string): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (slide) {
      slide.masterUri = masterUri;
      slide.masterName = undefined;
      slide.themeUri = undefined;
    }
  });
}

export function setPresentationSlideTheme(editor: OfficeEditor<PresentationDocument>, slideIndex: number, themeUri: string): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (slide) {
      slide.themeUri = themeUri;
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

export function setPresentationTimingNodes(editor: OfficeEditor<PresentationDocument>, slideIndex: number, nodes: Array<{ nodeType: string; presetClass?: string; presetId?: string; id?: string; duration?: string; repeatDuration?: string; repeatCount?: string; restart?: string; fill?: string; autoReverse?: boolean; acceleration?: string; deceleration?: string; triggerEvent?: string; triggerDelay?: string; endTriggerEvent?: string; endTriggerDelay?: string; targetShapeId?: string }>): PresentationDocument {
  return editor.transaction((draft) => {
    const slide = draft.slides[slideIndex];
    if (slide) {
      slide.timing = { nodeCount: nodes.length, nodes: [...nodes] };
    }
  });
}

export function setPresentationSize(editor: OfficeEditor<PresentationDocument>, size: { cx: number; cy: number }): PresentationDocument {
  return editor.transaction((draft) => {
    draft.size = { ...size };
  });
}

function ensurePresentationNotesPart(document: PresentationDocument, slideIndex: number): string | undefined {
  const slide = document.slides[slideIndex];
  if (!slide) {
    return undefined;
  }

  const existingUri = slide.notesUri;
  if (existingUri && document.packageGraph.parts[existingUri]) {
    return existingUri;
  }

  const notesUri = nextNotesUri(document.packageGraph.parts, slideIndex);
  const slideRelationships = relationshipsFor(document.packageGraph, slide.uri);
  const notesRelationship = slideRelationships.find((relationship) => relationship.type.includes('/notesSlide'));
  const relationshipId = notesRelationship?.id ?? nextRelationshipId(slideRelationships);

  updatePackagePartText(
    document.packageGraph,
    notesUri,
    `<?xml version="1.0" encoding="UTF-8"?>\n<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t></a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`,
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
  );
  upsertRelationship(document.packageGraph, slide.uri, {
    id: relationshipId,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    target: relativeRelationshipTarget(slide.uri, notesUri),
    targetMode: 'Internal'
  });

  return notesUri;
}

function ensurePresentationCommentsPart(document: PresentationDocument, slideIndex: number): string | undefined {
  const slide = document.slides[slideIndex];
  if (!slide) {
    return undefined;
  }

  const slideRelationships = relationshipsFor(document.packageGraph, slide.uri);
  const commentsRelationship = slideRelationships.find((relationship) => relationship.type.includes('/comments'));
  if (commentsRelationship?.resolvedTarget && document.packageGraph.parts[commentsRelationship.resolvedTarget]) {
    return commentsRelationship.resolvedTarget;
  }

  const commentsUri = nextCommentsUri(document.packageGraph.parts, slideIndex);
  const relationshipId = commentsRelationship?.id ?? nextRelationshipId(slideRelationships);
  updatePackagePartText(
    document.packageGraph,
    commentsUri,
    `<?xml version="1.0" encoding="UTF-8"?>\n<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:cmLst>`,
    'application/vnd.openxmlformats-officedocument.presentationml.comment+xml'
  );
  upsertRelationship(document.packageGraph, slide.uri, {
    id: relationshipId,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    target: relativeRelationshipTarget(slide.uri, commentsUri),
    targetMode: 'Internal'
  });

  return commentsUri;
}

function nextNotesUri(parts: PresentationDocument['packageGraph']['parts'], slideIndex: number): string {
  let candidateIndex = slideIndex + 1;
  let candidate = `/ppt/notesSlides/notesSlide${candidateIndex}.xml`;
  while (parts[candidate]) {
    candidateIndex += 1;
    candidate = `/ppt/notesSlides/notesSlide${candidateIndex}.xml`;
  }
  return candidate;
}

function nextCommentsUri(parts: PresentationDocument['packageGraph']['parts'], slideIndex: number): string {
  let candidateIndex = slideIndex + 1;
  let candidate = `/ppt/comments/comment${candidateIndex}.xml`;
  while (parts[candidate]) {
    candidateIndex += 1;
    candidate = `/ppt/comments/comment${candidateIndex}.xml`;
  }
  return candidate;
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

function relativeRelationshipTarget(sourceUri: string, targetUri: string): string {
  const sourceSegments = sourceUri.replace(/^\//, '').split('/');
  sourceSegments.pop();
  const targetSegments = targetUri.replace(/^\//, '').split('/');

  while (sourceSegments.length > 0 && targetSegments.length > 0 && sourceSegments[0] === targetSegments[0]) {
    sourceSegments.shift();
    targetSegments.shift();
  }

  return `${sourceSegments.map(() => '..').join('/')}${sourceSegments.length ? '/' : ''}${targetSegments.join('/')}`;
}
