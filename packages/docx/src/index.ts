export type {
  DocxComment,
  DocxHeaderFooterReference,
  DocxDocument,
  DocxSection,
  DocxSectionPageMargins,
  DocxSectionPageSize,
  DocxStyle,
  DocxParagraph,
  DocxRevision,
  DocxParagraphNumbering,
  DocxRun,
  DocxNumbering,
  DocxNumberingLevel,
  DocxNumberingInstance,
  DocxAbstractNumbering,
  DocxStory,
  DocxTable,
  DocxTableCell,
  DocxTableRow
} from './model';
export { parseDocx, resolveDocxNumbering, resolveDocxStyle } from './parser';
