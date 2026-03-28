export { createOfficeEditor } from './core';
export { replaceDocxParagraphText, replaceDocxStoryParagraphText, setDocxTableCellText, setDocxCommentText } from './docx';
export { setPresentationCommentText, setPresentationNotesText, setPresentationShapeText, setPresentationShapeTransform, setPresentationTransition } from './pptx';
export { insertWorkbookRow, setWorkbookCellValue, setWorkbookDefinedNameReference, setWorksheetCommentText, setWorksheetFrozenPane, setWorksheetMergedRanges, setWorksheetTableRange } from './xlsx';
export type { EditableOfficeDocument, OfficeEditor } from './types';
