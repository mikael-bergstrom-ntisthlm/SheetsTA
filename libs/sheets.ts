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

    spreadsheet.toast("Working on " + sheetName);

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.setFrozenRows(1);
    }
    else if (clear) {
      sheet.clear();
    }

    return sheet;
  }

  export function SetSheetSize(sheet: GoogleAppsScript.Spreadsheet.Sheet, targetWidth: number, targetHeight: number) {
    const currentHeight = sheet.getMaxRows();
    const currentWidth = sheet.getMaxColumns();

    // -- HEIGHT
    if (currentHeight < targetHeight) {
      AddEmptyRows(sheet, targetHeight - currentHeight);
    } else if (currentHeight > targetHeight) {
      sheet.deleteRows(
        targetHeight,
        currentHeight - targetHeight
      );
    }

    // -- WIDTH
    if (currentWidth < targetWidth) {
      AddEmptyColumns(sheet, targetWidth - currentWidth);
    } else if (currentWidth > targetWidth) {
      sheet.deleteColumns(
        targetWidth,
        currentWidth - targetWidth
      )
    }
  }

  function AddEmptyRows(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: number) {
    let emptyData: string[][] = [];
    for (let i = 0; i < rows; i++) {
      emptyData.push([""]);
    }

    const lastRow = sheet.getMaxRows();
    sheet.getRange(lastRow + 1, 1, rows, emptyData[0].length).setValues(emptyData);
  }

  function AddEmptyColumns(sheet: GoogleAppsScript.Spreadsheet.Sheet, cols: number) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      cols
    )
  }

  export function ClearSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet)
  {
    const height = sheet.getMaxRows();
    const width = sheet.getMaxColumns();

    sheet.clear();
    sheet.getFilter()?.remove();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    sheet.showColumns(1, width);
    
    sheet.getRange(1, 1,
      height,
      width)
      .removeCheckboxes()
      .setDataValidation(null)
      .getMergedRanges().forEach(mergedRange => mergedRange.breakApart());
  }

}