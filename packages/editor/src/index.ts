export { createOfficeEditor } from './core';
export { replaceDocxParagraphText, setDocxCommentText } from './docx';
export { setPresentationCommentText, setPresentationNotesText, setPresentationShapeText, setPresentationShapeTransform } from './pptx';
export { insertWorkbookRow, setWorkbookCellValue, setWorksheetCommentText, setWorksheetTableRange } from './xlsx';
export type { EditableOfficeDocument, OfficeEditor } from './types';
