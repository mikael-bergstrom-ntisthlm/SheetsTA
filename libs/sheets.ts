namespace SheetsUtilsTA {
  type RowProcessor = (row: any[]) => string[];

  export function ProcessCurrentRange(processor: RowProcessor) {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = sheet.getActiveRange();
    if (!range) return;

    const values = range.getValues();

    const colStart = range.getColumn();
    const rowStart = range.getRow();

    for (let rNum = 0; rNum < values.length; rNum++) {
      const row = values[rNum];
      const result: string[] = processor(row);

      if (result.length == 0) continue;

      let targetCells = sheet.getRange(rowStart + rNum, colStart + row.length, 1, result.length);
      targetCells.setValues([result]);
    }
  }
}