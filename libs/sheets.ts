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

    let maxWidth = values[0].length;
    values.forEach(row => { maxWidth = Math.max(maxWidth, row.length) });

    let targetRange = origo?.offset(0, 0, values.length, maxWidth);
    targetRange?.setValues(values);
  }

  export function CreateOrGetSheet(
    sheetName: string,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, clear: boolean): GoogleAppsScript.Spreadsheet.Sheet {

    spreadsheet.toast("Updating " + sheetName);

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.setFrozenRows(1);
    }
    else if (clear){
      sheet.clear();
    }

    return sheet;
  }

  export function GetColumnNum(columnHeading: string, sheet: GoogleAppsScript.Spreadsheet.Sheet, headingRowNumber: number): number {

    const headingRow = sheet.getRange(headingRowNumber, 1, 1, sheet.getMaxColumns())
    const headingRowValues = headingRow.getValues()[0];

    for (let i = 0; i < headingRowValues.length; i++) {
      if (headingRowValues[i] === columnHeading) return i + 1;
    }

    return -1;
  }
}