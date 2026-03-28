export type {
  DocxComment,
  DocxDocument,
  DocxStyle,
  DocxParagraph,
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
