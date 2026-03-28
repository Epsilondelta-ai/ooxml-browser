export { createOfficeEditor } from './core';
export { replaceDocxParagraphText, replaceDocxStoryParagraphText, setDocxSectionLayout, setDocxTableCellText, setDocxCommentAuthor, setDocxCommentText } from './docx';
export { setPresentationCommentAuthor, setPresentationCommentText, setPresentationNotesText, setPresentationShapeText, setPresentationShapeTransform, setPresentationTimingNodes, setPresentationTransition } from './pptx';
export { insertWorkbookRow, setWorkbookCellValue, setWorkbookDefinedNameReference, setWorksheetCommentAuthor, setWorksheetCommentText, setWorksheetFrozenPane, setWorksheetMergedRanges, setWorksheetTableName, setWorksheetTableRange } from './xlsx';
export type { EditableOfficeDocument, OfficeEditor } from './types';
