import { LibRubrics } from "../libs/rubrics";

export namespace PageGradingOverview {

  const _GradingOverviewSheetName = "OVERVIEW";

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

  export function GetGradingOverviewSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {
    
    return spreadsheet.getSheetByName(_GradingOverviewSheetName);
  }

  /**
   * Get the rubrics from the overview sheet
   * @param spreadsheet 
   * @returns {Rubric[]} an array of rubrics
   */
  export function GetRubrics(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet): LibRubrics.Rubric[] {

    // -- HEADER BLOCK VALUES RETRIEVAL
    let headerBlock = gradingOverviewSheet.getRange(
      1, gradingOverviewSheet.getFrozenColumns() + 1,
      gradingOverviewSheet.getFrozenRows(),
      gradingOverviewSheet.getLastColumn()
    );

    let headerValues = headerBlock?.getValues();

    if (!headerValues || headerValues?.length == 0) return [];

    // -- READ VALUES INTO DIFFERENT ROWS
    let rubricTitleRow = headerValues[0]; // TODO: Extract magic numbers
    let activeRow = headerValues[1]
    let gradeRow = headerValues[2];
    let shortformRow = headerValues[3]
    let criteriaRow = headerValues[4];

    let rubrics: LibRubrics.Rubric[] = [];
    let currentRubric: LibRubrics.Rubric | undefined = undefined;

    // Go through all columns of the rubric title row
    for (let i = 0; i < rubricTitleRow.length; i++) {
      // Detect rubric start
      if (rubricTitleRow[i] !== "") {
        currentRubric = {
          criteria: [],
          columnNumber: gradingOverviewSheet.getFrozenColumns() + i,
          name: rubricTitleRow[i]
        }
        rubrics.push(currentRubric);
      }

      // Detect criteria
      if (gradeRow[i] !== "" && currentRubric) {
        currentRubric.criteria.push(
          {
            name: criteriaRow[i],
            shortform: shortformRow[i],
            active: activeRow[i],
            grade: gradeRow[i],
            columnNumber: gradingOverviewSheet.getFrozenColumns() + i
          }
        )
      }
    }

    return rubrics;
  }

  /**
   * Retrieve basic info of all students from a grading overview sheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet whose overview to get students from
   * @returns {StudentData[]} an array of student data
   */
  export function GetStudentsData(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const studentValues = gradingOverviewSheet.getRange(
      _RowDataStart, 1,
      gradingOverviewSheet.getLastRow() - _RowDataStart + 1,
      _ColDataStart
    ).getValues()

    const studentsData: StudentData[] = [];

    studentValues.forEach(row => {
      // Skip empties
      if (row[_ColUserId - 1].length === 0) return;

      studentsData.push({
        id: row[_ColUserId - 1],
        name: row[_ColName - 1],
        surname: row[_ColSurname - 1],
        email: row[_ColEmail - 1]
      });
    });

    return studentsData;
  }


  /**
   * Get the details of a single user from the overview sheet
   * @param {string} userID - The user to get the details of
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet containing the overview sheet
   * @returns {StudentData} the data of the student
   */
  export function GetStudentData(userID: string, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): StudentData | null {

    const gradingOverviewSheet = spreadsheet.getSheetByName(_GradingOverviewSheetName);
    if (!gradingOverviewSheet) return null;

    const studentsData = GetStudentsData(gradingOverviewSheet);

    let rowNum = studentsData.findIndex(student => student.id === userID);
    if (rowNum < 0) return null;

    const student = studentsData[rowNum];
    student.dataRange = gradingOverviewSheet.getRange(_RowDataStart + rowNum, 1, 1, gradingOverviewSheet.getMaxColumns())

    return student;
  }


  export interface StudentData {
    id: string,
    name: string,
    surname: string,
    email: string,
    dataRange?: GoogleAppsScript.Spreadsheet.Range
  }
}