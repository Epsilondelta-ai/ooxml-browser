export { createOfficeEditor } from './core';
export { replaceDocxParagraphText, replaceDocxStoryParagraphText, setDocxParagraphNumbering, setDocxParagraphRunStyle, setDocxParagraphStyle, setDocxSectionLayout, setDocxTableCellText, setDocxCommentAuthor, setDocxCommentText } from './docx';
export { setPresentationCommentAuthor, setPresentationCommentText, setPresentationNotesText, setPresentationShapeName, setPresentationShapePlaceholderType, setPresentationShapeText, setPresentationShapeTransform, setPresentationSize, setPresentationTimingNodes, setPresentationTransition } from './pptx';
export { insertWorkbookRow, setWorkbookCellFormula, setWorkbookCellValue, setWorkbookDefinedNameReference, setWorkbookSheetName, setWorksheetCommentAuthor, setWorksheetCommentText, setWorksheetFrozenPane, setWorksheetMergedRanges, setWorksheetTableName, setWorksheetTableRange } from './xlsx';
export type { EditableOfficeDocument, OfficeEditor } from './types';
