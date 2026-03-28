export type {
  DocxComment,
  DocxDocument,
  DocxStyle,
  DocxParagraph,
  DocxRun,
  DocxStory,
  DocxTable,
  DocxTableCell,
  DocxTableRow
} from './model';
export { parseDocx, resolveDocxStyle } from './parser';
