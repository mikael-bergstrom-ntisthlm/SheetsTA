namespace MasterGradingSheetTA {

  interface StudentData {
    id: string,
    name: string,
    surname: string,
    email: string,
    dataRange?: GoogleAppsScript.Spreadsheet.Range
  }

  const _ColClassroomID = 1;
  const _ColCourseID = 2;
  const _ColName = 3;
  const _ColSurname = 4;
  const _ColEmail = 5;
  const _ColUserId = 6;

  const _RowRubricTitle = 1;
  const _RowCriteriaActive = 2;
  const _RowGrade = 3;
  const _RowTag = 4;
  const _RowHeading = 5;
  const _RowDataStart = 6;

  // TODO: Rebuild to generate arbitrary data block w/ filter
  export function GetStudentData(userID: string,
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): StudentData | null {

    const headingRowNumber = masterGradingSheet.getFrozenRows();

    // Get the ID column
    const idColumnRange = masterGradingSheet.getRange(headingRowNumber + 1, _ColUserId, masterGradingSheet.getMaxRows());

    let IDs = idColumnRange.getValues().map(cell => cell[0]).filter(cell => cell.length > 0);
    let rowNum = IDs.findIndex(id => id === userID);

    if (rowNum < 0) return null;

    const studentRange = masterGradingSheet.getRange(headingRowNumber + rowNum + 1, 1, 1, masterGradingSheet.getMaxColumns())
    const studentValues = studentRange.getValues()[0];

    return {
      id: studentValues[_ColUserId],
      name: studentValues[_ColName],
      surname: studentValues[_ColSurname],
      email: studentValues[_ColEmail],
      dataRange: studentRange
    }
  }

  export function GetStudentNameIds(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const headingRowNumber = masterGradingSheet.getFrozenRows();
    const dataHeight = masterGradingSheet.getLastRow() - headingRowNumber;

    let studentIds = GetContentsOfColumn(_ColUserId, dataHeight, masterGradingSheet);
    let studentSurnames = GetContentsOfColumn(_ColSurname, dataHeight, masterGradingSheet);
    let studentNames = GetContentsOfColumn(_ColName, dataHeight, masterGradingSheet);

    const studentList: string[] = [];

    for (let i = 0; i < studentIds.length; i++) {
      if (studentIds[i] === "") continue;

      studentList.push(
        studentNames[i] + " " + studentSurnames[i] + " | " + studentIds[i]
      );
    }

    return studentList;
  }

  function GetContentsOfColumn(colNum: number, dataHeight: number, sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {

    let data = sheet.getRange(_RowDataStart, colNum, dataHeight)
      .getValues()
      .map(item => item[0]);

    return data;
  }
}