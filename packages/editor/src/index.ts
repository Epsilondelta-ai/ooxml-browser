export { createOfficeEditor } from './core';
export { replaceDocxParagraphText, replaceDocxStoryParagraphText, setDocxTableCellText, setDocxCommentAuthor, setDocxCommentText } from './docx';
export { setPresentationCommentText, setPresentationNotesText, setPresentationShapeText, setPresentationShapeTransform, setPresentationTimingNodes, setPresentationTransition } from './pptx';
export { insertWorkbookRow, setWorkbookCellValue, setWorkbookDefinedNameReference, setWorksheetCommentText, setWorksheetFrozenPane, setWorksheetMergedRanges, setWorksheetTableRange } from './xlsx';
export type { EditableOfficeDocument, OfficeEditor } from './types';
