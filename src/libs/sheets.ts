export namespace LibGSheets {

  /**
   * Find a sheet and return it. Create it if it doesn't exist
   * @param {string} sheetName - The name of the sgeet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The Sheets document to look in
   * @param {boolean} clear - Empty the sheet before returning it
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet that was found or created
   */
  export function CreateOrGetSheet(
    sheetName: string,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, clear: boolean): GoogleAppsScript.Spreadsheet.Sheet {

    // TODO: Centralize logging, toasts & alerts
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

  
  /**
   * Adjust the size of a sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - the sheet to resize
   * @param {number} targetWidth - The target width
   * @param {number} targetHeight - The target height
   */
  export function SetSheetSize(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    targetWidth: number,
    targetHeight: number) {

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


  /**
   * Add empty rows to a sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet
   * @param {number} rows - The number of new rows
   */
  function AddEmptyRows(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: number) {
    let emptyData: string[][] = [];
    for (let i = 0; i < rows; i++) {
      emptyData.push([""]);
    }

    const lastRow = sheet.getMaxRows();
    sheet.getRange(lastRow + 1, 1, rows, emptyData[0].length).setValues(emptyData);
  }

  /**
   * Add empty columns to a sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet
   * @param {number} cols - The number of new columns
   */
  function AddEmptyColumns(sheet: GoogleAppsScript.Spreadsheet.Sheet, cols: number) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      cols
    )
  }


  /**
   * Insert some values into a Sheet, beginning at a specific origo
   * @param {string[][]} values - The values to insert
   * @param {GoogleAppsScript.Spreadsheet.Range} origo - The origo to begin inserting at
   */
  export function InsertValuesAt(
    values: string[][],
    origo: GoogleAppsScript.Spreadsheet.Range) {

    let maxWidth = values[0].length;
    values.forEach(row => { maxWidth = Math.max(maxWidth, row.length) });

    let targetRange = origo?.offset(0, 0, values.length, maxWidth);
    targetRange?.setValues(values);
  }


  /**
   * Completely clear a sheet of contents, formatting, frozen columns/rows etc
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet
   */
  export function ClearSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
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


  /**
   * Go through the currently selected range, run all rows through the given
   *   row processor function, and insert the results to the right of the 
   *   original rows
   * @param {RowProcessor} processor - A function that takes an any[] array (a 
   *   row) as parameter and returns a string array
   * @returns
   */
  export function ProcessCurrentRange(processor: RowProcessor) {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = sheet.getActiveRange();
    if (!range) return;

    const values = range.getValues();

    const colStart = range.getColumn();
    const rowStart = range.getRow();

    // Go through all rows
    for (let rNum = 0; rNum < values.length; rNum++) {
      // Get row values
      const row = values[rNum];
      // Get resulting string
      const result: string[] = processor(row);

      if (result.length == 0) continue;

      let targetCells = sheet.getRange(
        rowStart + rNum, // current row's index
        colStart + row.length, // column to the right of the range's last col
        1, // One row
        result.length // As many columns as needed
      );

      targetCells.setValues([result]);
    }
  }

  // A RowProcessor is a function that takes an any[] array (a row) as parameter
  //  and returns a string array
  type RowProcessor = (row: any[]) => string[];
}