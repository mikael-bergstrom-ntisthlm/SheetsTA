namespace MasterGradingSheetTA {
  export function GetStudentRange(userID: string, masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): GoogleAppsScript.Spreadsheet.Range | null {
    const headingRowNumber = masterGradingSheet.getFrozenRows();

    // Get the ID column
    let colnumId = SheetsTA.GetColumnNum("UserID", masterGradingSheet, headingRowNumber) // 6
    const idColumnRange = masterGradingSheet.getRange(headingRowNumber + 1, colnumId, masterGradingSheet.getMaxRows());

    // Find the right row
    let IDs = idColumnRange.getValues().map(cell => cell[0]).filter(cell => cell.length > 0);
    let rowNum = IDs.findIndex(id => id === userID);

    if (rowNum < 0) return null;

    const studentRange = masterGradingSheet.getRange(headingRowNumber + rowNum + 1, 1, 1, masterGradingSheet.getMaxColumns())

    return studentRange;
  }

  export function GetStudentNameIds(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const headingRowNumber = masterGradingSheet.getFrozenRows();
    const dataHeight = masterGradingSheet.getMaxRows() - headingRowNumber;

    let studentIds = GetContentsOfColumn("userid", headingRowNumber, dataHeight, masterGradingSheet);
    let studentSurnames = GetContentsOfColumn("surname", headingRowNumber, dataHeight, masterGradingSheet);
    let studentNames = GetContentsOfColumn("name", headingRowNumber, dataHeight, masterGradingSheet);

    const studentList: string[] = [];

    for (let i = 0; i < studentIds.length; i++) {
      if (studentIds[i] === "") continue;

      studentList.push(
        studentNames[i] + " " + studentSurnames[i] + " | " + studentIds[i]
      );
    }

    return studentList;
  }

  function GetContentsOfColumn(columnName: string, headingRowNumber: number, dataHeight: number, sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {

    let colId = SheetsTA.GetColumnNum(columnName, sheet, headingRowNumber);
    if (colId < 0) Browser.msgBox(`Column "${columnName}" not found`);

    let data = sheet.getRange(headingRowNumber + 1, colId, dataHeight)
      .getValues()
      .map(item => item[0]);

    return data;
  }
}