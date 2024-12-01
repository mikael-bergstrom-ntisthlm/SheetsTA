namespace SheetsTA {
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

  export function InsertValuesAt(values: string[][], origo: GoogleAppsScript.Spreadsheet.Range) {
    let targetRange = origo?.offset(0, 0, values.length, values[0].length);
    targetRange?.setValues(values);
  }
}