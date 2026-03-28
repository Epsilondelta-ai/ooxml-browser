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

export function setWorkbookCellFormula(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, formula: string, value: string): XlsxWorkbook {
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

      cell.formula = formula;
      cell.value = value;
      cell.type = 'n';
      return;
    }
  });
}

export function setWorkbookCellStyle(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, styleIndex: number | undefined): XlsxWorkbook {
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

      cell.styleIndex = styleIndex;
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

export function setWorksheetCommentAuthor(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, author: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    const comment = sheet.comments.find((entry) => entry.reference === reference);
    if (comment) {
      comment.author = author;
    }
  });
}

export function removeWorksheetComment(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    sheet.comments = sheet.comments.filter((entry) => entry.reference !== reference);
  });
}

export function upsertWorksheetComment(editor: OfficeEditor<XlsxWorkbook>, sheetName: string, reference: string, text: string, author?: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === sheetName);
    if (!sheet) {
      return;
    }

    const existing = sheet.comments.find((entry) => entry.reference === reference);
    if (existing) {
      existing.text = text;
      if (author !== undefined) {
        existing.author = author;
      }
      return;
    }

    sheet.comments.push({ reference, text, author });
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

export function upsertWorkbookDefinedName(editor: OfficeEditor<XlsxWorkbook>, name: string, reference: string, scopeSheetId?: number): XlsxWorkbook {
  return editor.transaction((draft) => {
    const definedName = draft.definedNames.find((entry) => entry.name === name && entry.scopeSheetId === scopeSheetId);
    if (definedName) {
      definedName.reference = reference;
      return;
    }

    draft.definedNames.push({ name, reference, scopeSheetId });
  });
}

export function setWorkbookSheetName(editor: OfficeEditor<XlsxWorkbook>, currentName: string, nextName: string): XlsxWorkbook {
  return editor.transaction((draft) => {
    const sheet = draft.sheets.find((entry) => entry.name === currentName);
    if (!sheet) {
      return;
    }

    sheet.name = nextName;
    const renameSheetRef = createSheetReferenceRewriter(currentName, nextName);

    for (const workbookSheet of draft.sheets) {
      for (const row of workbookSheet.rows) {
        for (const cell of row.cells) {
          if (cell.formula) {
            cell.formula = renameSheetRef(cell.formula);
          }
        }
      }
    }

    draft.definedNames = draft.definedNames.map((definedName) => ({
      ...definedName,
      reference: renameSheetRef(definedName.reference)
    }));
  });
}

function createSheetReferenceRewriter(currentName: string, nextName: string): (value: string) => string {
  const escapedCurrent = currentName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const quotedPattern = new RegExp(`'${escapedCurrent}'!`, 'g');
  const barePattern = new RegExp(`(^|[^A-Za-z0-9_])${escapedCurrent}!`, 'g');

  return (value: string) => value
    .replace(quotedPattern, `'${nextName}'!`)
    .replace(barePattern, (_match, prefix) => `${prefix}${nextName}!`);
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
