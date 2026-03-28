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
