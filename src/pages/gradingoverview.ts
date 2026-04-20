import { LibGClassroom } from "../libs/classroom.js";
import { LibConfig } from "../libs/config.js";
import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js";
import { PageRubrics } from "./rubrics.js";

export namespace PageGradingOverview {

  const _GradingOverviewSheetName = "OVERVIEW";

  const _ColClassroomID = 1;
  const _ColCourseID = 2;
  const _ColName = 3;
  const _ColSurname = 4;
  const _ColEmail = 5;
  const _ColUserId = 6;
  const _ColFullName = 7;
  const _ColOutput = 8;

  const _RowRubricTitle = 1;
  const _RowCriteriaActive = 2;
  const _RowGrade = 3;
  const _RowTag = 4;
  const _RowHeading = 5;
  const _RowDataStart = 6;

  export function GetDefaultGradingOverviewSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_GradingOverviewSheetName);
  }

  export namespace Setup {
    export function Setup(
      spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
      config: LibConfig.Config
    ) {

      const gradingOverviewSheet = LibGSheets.CreateOrGetSheet(
        _GradingOverviewSheetName,
        spreadsheet, true
      );
      LibGSheets.ClearSheet(gradingOverviewSheet);

      const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);
      if (!rubricsSheet) return;

      // Get rubrics from _RUBRICS
      let rubrics = PageRubrics.GetRubrics(rubricsSheet);

      // Initialize some values
      let allCriteria = rubrics.flatMap(rubric => rubric.criteria);
      let highestCriteriaColId = Math.max(...allCriteria.map(criteria => criteria.columnNumber));

      let totalWidth = LibGClassroom.rosterHeaders.length + 2
        + allCriteria.length
        + rubrics.length * 2
        + 2; // margin
      LibGSheets.SetSheetWidth(gradingOverviewSheet, totalWidth);

      // Setup headers
      SetupRosterHeader(gradingOverviewSheet);
      let startColumn = gradingOverviewSheet.getLastColumn() + 1;
      SetupRubricHeader(gradingOverviewSheet, startColumn, highestCriteriaColId, rubrics);

      // Setup data areas
      SetupRosterArea(gradingOverviewSheet);

      // -- Set overall visuals
      FormatHeader(gradingOverviewSheet, startColumn, highestCriteriaColId);
    }

    function SetupRosterArea(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet) {
      let fullNameRange = gradingOverviewSheet.getRange(
        _RowDataStart, _ColFullName,
        gradingOverviewSheet.getMaxRows() - _RowDataStart
      );

      fullNameRange.setFormula(`=${String.fromCharCode(64 + _ColSurname)}${_RowDataStart} & " " & ${String.fromCharCode(64 + _ColName)}${_RowDataStart}`);

      fullNameRange.setBackground("#d9d9d9");
    }

    export function UpdateActiveCriteriaFromTemplate(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
      const gradingOverviewSheet = GetDefaultGradingOverviewSheet(spreadsheet);
      const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);

      if (!gradingOverviewSheet || !rubricsSheet) return;

      let rubrics = PageRubrics.GetRubrics(rubricsSheet);

      let startColumn = gradingOverviewSheet.getFrozenColumns();
      let allCriteria = rubrics.flatMap(rubric => rubric.criteria);
      let highestCriteriaColId = Math.max(...allCriteria.map(criteria => criteria.columnNumber));

      // Get the range we need
      let rubricHeaderRange = gradingOverviewSheet.getRange(
        _RowCriteriaActive, startColumn,
        1, startColumn + highestCriteriaColId
      );
      let rubricHeaderRangeValues = rubricHeaderRange.getValues();


      allCriteria.forEach(criteria => {
        rubricHeaderRangeValues[0][criteria.columnNumber] = criteria.active;
      });


      rubricHeaderRange.setValues(rubricHeaderRangeValues);
    }

    export function UpdateActiveCriteriaToTemplate() {
      // TODO: Implement
    }

    function SetupRosterHeader(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet) {

      // Prepare headers
      let rosterHeaders = LibGClassroom.rosterHeaders;
      rosterHeaders.push("Full name");
      rosterHeaders.push("Output");

      // Freeze rows & columns to create quadrants
      gradingOverviewSheet.setFrozenColumns(rosterHeaders.length);
      gradingOverviewSheet.setFrozenRows(_RowDataStart - 1);

      // Setup roster headings (top-left quadrant)
      let headerRange = gradingOverviewSheet.getRange(_RowTag, 1, 2, rosterHeaders.length);
      let headerRangeValues = headerRange.getValues();

      // TODO: Use the consts for col-numbers
      headerRangeValues[0] = rosterHeaders.map(
        v => LibRubrics.GetSafeTagName(v)
      );
      headerRangeValues[1] = rosterHeaders;

      headerRange.setValues(headerRangeValues);

      gradingOverviewSheet.setColumnWidth(_ColName, 150);
      gradingOverviewSheet.setColumnWidth(_ColSurname, 150);
      gradingOverviewSheet.setColumnWidth(_ColClassroomID, 150);
      gradingOverviewSheet.setColumnWidth(_ColFullName, 150);
      gradingOverviewSheet.setColumnWidth(_ColOutput, 50);

      gradingOverviewSheet.hideColumns(_ColUserId);
      gradingOverviewSheet.hideColumns(_ColCourseID);
      gradingOverviewSheet.hideColumns(_ColEmail);
    }

    function SetupRubricHeader(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet, startColumn: number, highestCriteriaColId: number, rubrics: LibRubrics.Rubric[]) {
      // Get the range we need
      let rubricHeaderRange = gradingOverviewSheet.getRange(
        1, startColumn,
        5, highestCriteriaColId + 3 // + for extra cols after rubrics (for results export)
      );
      let rubricHeaderRangeValues = rubricHeaderRange.getValues();

      // Go through the rubrics
      rubrics.forEach(rubric => {
        // Set rubric heading
        rubricHeaderRangeValues[_RowRubricTitle - 1][rubric.columnNumber - 1]
          = rubric.name;

        // Insert criteria info
        rubric.criteria.forEach(criteria => {
          rubricHeaderRangeValues[_RowHeading - 1][criteria.columnNumber - 1]
            = criteria.name;
          rubricHeaderRangeValues[_RowTag - 1][criteria.columnNumber - 1]
            = criteria.tag;
          rubricHeaderRangeValues[_RowGrade - 1][criteria.columnNumber - 1]
            = criteria.grade;
          rubricHeaderRangeValues[_RowCriteriaActive - 1][criteria.columnNumber - 1]
            = criteria.active;
        });

        let lastColumnOfRubric = Math.max(...rubric.criteria.map(criteria => criteria.columnNumber));

        // -- Setup Grade column for rubric
        rubricHeaderRangeValues[_RowTag - 1][lastColumnOfRubric]
          = rubric.gradeTag;
        rubricHeaderRangeValues[_RowHeading - 1][lastColumnOfRubric]
          = "Grade";
        rubricHeaderRangeValues[_RowCriteriaActive - 1][lastColumnOfRubric]
          = true;
        FormatRubricSingleHeader(gradingOverviewSheet, startColumn, rubric);
      });

      // Response doc header
      rubricHeaderRangeValues[_RowHeading - 1][rubricHeaderRangeValues[_RowHeading - 1].length - 1] = "RESPONSE";
      rubricHeaderRangeValues[_RowTag - 1][rubricHeaderRangeValues[_RowTag - 1].length - 1] = "responsedoc";

      rubricHeaderRange.setValues(rubricHeaderRangeValues);
    }

    function FormatRubricSingleHeader(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet, startColumn: number, rubric: LibRubrics.Rubric) {
      // -- Rubric titles visuals
      let rubricTitleRange = gradingOverviewSheet.getRange(
        _RowRubricTitle, startColumn + rubric.columnNumber - 1,
        1, rubric.criteria.length + 1);
      rubricTitleRange.merge();
      rubricTitleRange.setFontWeight("bold");

      // -- Criteria active checkboxes
      let criteriaCheckboxRange = gradingOverviewSheet.getRange(
        _RowCriteriaActive, startColumn + rubric.columnNumber - 1,
        1, rubric.criteria.length + 1);
      criteriaCheckboxRange.insertCheckboxes();

      // -- Column widths
      gradingOverviewSheet.setColumnWidths(
        startColumn + rubric.columnNumber, rubric.criteria.length, 100
      );
      gradingOverviewSheet.setColumnWidth(
        startColumn + rubric.columnNumber + rubric.criteria.length, 20
      );
    }

    function FormatHeader(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet, startColumn: number, highestCriteriaColId: number) {
      let lastCol = gradingOverviewSheet.getLastColumn();
      let headingRange = gradingOverviewSheet.getRange(
        _RowHeading, 1, 1,
        lastCol);
      let tagRange = gradingOverviewSheet.getRange(
        _RowTag, 1, 1,
        lastCol);

      headingRange.setFontWeight("bold");
      headingRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      tagRange.setFontSize(8);
      tagRange.setFontStyle("italic");
    }
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
    let rubricTitleRow = headerValues[_RowRubricTitle - 1];
    let activeRow = headerValues[_RowCriteriaActive - 1]
    let gradeRow = headerValues[_RowGrade - 1];
    let tagRow = headerValues[_RowTag - 1]
    let criteriaRow = headerValues[_RowHeading - 1];

    let rubrics: LibRubrics.Rubric[] = [];
    let currentRubric: LibRubrics.Rubric | undefined = undefined;

    // Go through all columns of the rubric title row
    for (let i = 0; i < rubricTitleRow.length; i++) {
      // Detect rubric start
      if (rubricTitleRow[i] !== "") {
        currentRubric = {
          criteria: [],
          columnNumber: gradingOverviewSheet.getFrozenColumns() + i,
          name: rubricTitleRow[i],
          gradeTag: LibRubrics.GetSafeTagName(rubricTitleRow[i]) + "grade"
        }
        rubrics.push(currentRubric);
      }

      // Detect criteria
      if (gradeRow[i] !== "" && currentRubric) {
        currentRubric.criteria.push(
          {
            name: criteriaRow[i],
            tag: tagRow[i],
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
      gradingOverviewSheet.getFrozenColumns()
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
   * Get the details of a single user from the overview sheet, including data Range
   * @param {string} userID - The user to get the details of
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet containing the overview sheet
   * @returns {StudentData} the data of the student
   */
  export function GetStudentData(userID: string,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet): StudentData | null {

    const studentsData = GetStudentsData(gradingOverviewSheet);

    let rowNum = studentsData.findIndex(student => student.id === userID);
    if (rowNum < 0) return null;

    const student = studentsData[rowNum];
    student.dataRange = gradingOverviewSheet.getRange(_RowDataStart + rowNum, 1, 1, gradingOverviewSheet.getMaxColumns());
    return student;
  }


  export function InsertRubricData(userID: any, rubrics: LibRubrics.Rubric[]) {
    throw new Error("Function not implemented.");
  }

  /**
   * Get the details of a single user from the overview sheet, including rubrics
   * @param userID 
   * @param rubricsSheet 
   * @param gradingOverviewSheet 
   * @returns 
   */
  export function GetStudentDataRubrics(userID: string,
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet): StudentData | null {

    // TODO: Make this more precise
    const colDataStart = gradingOverviewSheet.getFrozenColumns() + 1;

    // Find the right student
    const studentsData = GetStudentsData(gradingOverviewSheet);

    let studentRowNum = studentsData.findIndex(student => student.id === userID);
    if (studentRowNum < 0) { Browser.msgBox("Student ID not found"); return null; };

    const student = studentsData[studentRowNum];

    // Get the student's data
    const studentDataValues = gradingOverviewSheet.getRange(
      _RowDataStart + studentRowNum, // Student's row
      colDataStart,
      1, // only one row
      gradingOverviewSheet.getMaxColumns() - colDataStart
    ).getValues();

    // Get the rubrics from the rubrics sheet
    student.rubricData = PageRubrics.GetRubrics(rubricsSheet);

    // Get the tags-row from the overview sheet
    const tagsValues = gradingOverviewSheet.getRange(
      _RowTag, colDataStart,
      1, gradingOverviewSheet.getMaxColumns() - colDataStart
    ).getValues();

    // Go through the rubrics
    student.rubricData.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // TODO: remove reliance on columnNumber; just find matching tag
        // Check if the tags match
        if (criterion.tag == tagsValues[0][criterion.columnNumber - 1]) {
          criterion.studentPassed =
            studentDataValues[0][criterion.columnNumber - 1] == "✔";

        } else {
          Browser.msgBox(`MISMATCH:\\n${criterion.tag} != ${tagsValues[0][criterion.columnNumber - 1]}`);
        }
      });

      // Find column of rubric's overall grade
      const gradeCol = rubric.criteria.reduce(
        (prev, current) => {
          return prev.columnNumber > current.columnNumber ? prev : current
        }
      ).columnNumber;

      // Save rubric grade
      rubric.studentGrade = studentDataValues[0][gradeCol];

    });

    return student;
  }

  export interface StudentData {
    id: string,
    name: string,
    surname: string,
    email: string,
    dataRange?: GoogleAppsScript.Spreadsheet.Range,
    rubricData?: LibRubrics.Rubric[]
  }
}