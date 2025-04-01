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
  const _ColDataStart = 7;

  const _RowRubricTitle = 1;
  const _RowCriteriaActive = 2;
  const _RowGrade = 3;
  const _RowTag = 4;
  const _RowHeading = 5;
  const _RowDataStart = 6;

  export function GetStudentData(userID: string,
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet
  ): StudentData | null {

    const studentsData = GetStudentsData(masterGradingSheet);

    let rowNum = studentsData.findIndex(student => student.id === userID);
    if (rowNum < 0) return null;
    
    const student = studentsData[rowNum];
    student.dataRange = masterGradingSheet.getRange(_RowDataStart + rowNum, 1, 1, masterGradingSheet.getMaxColumns())

    return student;
  }

  export function GetStudentsData(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const studentValues = masterGradingSheet.getRange(
      _RowDataStart, 1,
      masterGradingSheet.getLastRow() - _RowDataStart + 1,
      _ColDataStart
    ).getValues()

    const studentsData: StudentData[] = [];

    studentValues.forEach(row => {
      studentsData.push({
        id: row[_ColUserId - 1],
        name: row[_ColName - 1],
        surname: row[_ColSurname - 1],
        email: row[_ColEmail - 1]
      });
    });

    return studentsData;
  }
}