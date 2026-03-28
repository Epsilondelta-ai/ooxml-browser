import type { XlsxWorkbook } from '@ooxml/xlsx';

import type { OfficeEditor } from './types';

export function setWorkbookCellValue(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, value: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    for (const row of sheet.rows) {
      const cell = row.cells.find((entry) => entry.reference === reference);
      if (!cell) {
        continue;
      }

      cell.value = value;
      cell.type = Number.isNaN(Number(value)) ? 's' : 'n';
      cell.formula = undefined;
      return;
    }
  });
}


export function insertWorkbookRow(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, rowIndex: number): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    for (const row of sheet.rows) {
      if (row.index >= rowIndex) {
        row.index += 1;
      }

      for (const cell of row.cells) {
        cell.reference = shiftCellReferenceRow(cell.reference, rowIndex, 1);
        if (cell.formula) {
          cell.formula = shiftFormulaRowReferences(cell.formula, rowIndex, 1);
        }
      }
    }

    sheet.rows.sort((left, right) => left.index - right.index);
    sheet.rows.push({ index: rowIndex, cells: [] });
    sheet.rows.sort((left, right) => left.index - right.index);

    sheet.mergedRanges = sheet.mergedRanges.map((range) => shiftRangeRows(range, rowIndex, 1));
    if (sheet.frozenPane?.topLeftCell) {
      sheet.frozenPane.topLeftCell = shiftCellReferenceRow(sheet.frozenPane.topLeftCell, rowIndex, 1);
    }
    sheet.tables = sheet.tables.map((table) => ({
      ...table,
      range: shiftRangeRows(table.range, rowIndex, 1)
    }));
    sheet.comments = sheet.comments.map((comment) => ({
      ...comment,
      reference: shiftCellReferenceRow(comment.reference, rowIndex, 1)
    }));

    draft.definedNames = draft.definedNames.map((definedName) => ({
      ...definedName,
      reference: shiftReferenceRows(definedName.reference, rowIndex, 1)
    }));
  });
}

function shiftFormulaRowReferences(formula: string, threshold: number, delta: number): string {
  return formula.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (_match, column, row) => {
    const numericRow = Number(row);
    return `${column}${numericRow >= threshold ? numericRow + delta : numericRow}`;
  });
}

function shiftCellReferenceRow(reference: string, threshold: number, delta: number): string {
  return reference.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/, (_match, column, row) => {
    const numericRow = Number(row);
    return `${column}${numericRow >= threshold ? numericRow + delta : numericRow}`;
  });
}

function shiftRangeRows(range: string, threshold: number, delta: number): string {
  return range.split(':').map((entry) => shiftCellReferenceRow(entry, threshold, delta)).join(':');
}

function shiftReferenceRows(reference: string, threshold: number, delta: number): string {
  return reference.replace(/(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/g, (match) => shiftRangeRows(match, threshold, delta));
}


export function setWorksheetCommentText(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, text: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    const comment = sheet.comments.find((entry) => entry.reference === reference);
    if (comment) {
      comment.text = text;
    }
  });
}

export function setWorksheetTableRange(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, tableName: string, range: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    const table = sheet.tables.find((entry) => entry.name === tableName);
    if (table) {
      table.range = range;
    }
  });
}

export function setWorksheetTableName(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, tableName: string, nextName: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    const table = sheet.tables.find((entry) => entry.name === tableName);
    if (table) {
      table.name = nextName;
    }
  });
}

export function setWorkbookDefinedNameReference(editor: OfficeEditor<XlsxWorkbook>, name: string, reference: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const definedName = draft.definedNames.find((entry) => entry.name === name);
    if (definedName) {
      definedName.reference = reference;
    }
  });
}

export function setWorksheetFrozenPane(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, frozenPane: { xSplit?: number; ySplit?: number; topLeftCell?: string; state?: string } | undefined): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (sheet) {
      sheet.frozenPane = frozenPane;
    }
  });
}

export function setWorksheetMergedRanges(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, mergedRanges: string[]): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (sheet) {
      sheet.mergedRanges = [...mergedRanges];
    }
  });
}
