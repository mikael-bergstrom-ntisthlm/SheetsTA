import { LibGClassroom } from "../libs/classroom.js";
import { LibConfig } from "../libs/config.js";
import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js";
import { LibStudents } from "../libs/students.js";
import { PageResponse } from "./response.js";
import { PageRubrics } from "./rubrics.js";

//TODO: Implement "Name of assignment"
//TODO: Implement auto-adding roster
//TODO: Implement adding a filtered assignment column

export namespace PageGradingOverview {

  const _GradingOverviewSheetName = "OVERVIEW";

  const _ColClassroomID = 1;
  const _ColCourseID = 2;
  const _ColFullName = 7;
  const _ColOutput = 8;

  const _RowRubricTitle = 1;
  const _RowCriteriaActive = 2;
  const _RowGrade = 3;
  const _RowTag = 4;
  const _RowHeading = 5;
  const _RowDataStart = 6;

  export const studentColumnSetup: LibStudents.StudentColumnSetup = {
    colName: 3,
    colSurname: 4,
    colEmail: 5,
    colUserId: 6
  }

  export function GetDefaultGradingOverviewSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_GradingOverviewSheetName);
  }

  export namespace Setup {
    export function Setup(
      spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
      config: LibConfig.Config
    ) {
      // TODO: Add some sort of warning if there's already data

      // TODO: Implement automatic adding of a filter

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


      let totalWidth = LibGClassroom.rosterHeaders.length + 2
        + GetTotalWidthNeeded(rubrics) + 4;

      LibGSheets.SetSheetWidth(gradingOverviewSheet, totalWidth);

      // Setup headers
      SetupRosterHeader(gradingOverviewSheet);
      let startColumn = gradingOverviewSheet.getLastColumn() + 1;
      SetupRubricHeader(gradingOverviewSheet, startColumn, rubrics);

      // Setup data areas
      SetupRosterArea(gradingOverviewSheet);

      // -- Set overall visuals
      FormatHeader(gradingOverviewSheet);
    }

    function SetupRosterArea(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet) {
      let fullNameRange = gradingOverviewSheet.getRange(
        _RowDataStart, _ColFullName,
        gradingOverviewSheet.getMaxRows() - _RowDataStart
      );

      fullNameRange.setFormula(`=${String.fromCharCode(64 + studentColumnSetup.colSurname)}${_RowDataStart} & " " & ${String.fromCharCode(64 + studentColumnSetup.colName)}${_RowDataStart}`);

      fullNameRange.setBackground("#d9d9d9");
    }

    export function UpdateActiveCriteriaFromTemplate(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
      const gradingOverviewSheet = GetDefaultGradingOverviewSheet(spreadsheet);
      const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);

      if (!gradingOverviewSheet || !rubricsSheet) return;

      let rubrics = PageRubrics.GetRubrics(rubricsSheet);

      // Get the range we need
      let rubricHeaderRange = gradingOverviewSheet.getRange(
        _RowCriteriaActive, 1,
        1, gradingOverviewSheet.getLastColumn()
      );
      let rubricHeaderRangeValues = rubricHeaderRange.getValues();

      // -- Make a map of which column belongs to which tag
      const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

      // -- Go through all rubrics, insert checkmarks & grades
      rubrics.forEach(rubric => {
        rubric.criteria.forEach(criterion => {
          // Find the column with a matching tag
          const colNumber = tagColNumbers.get(criterion.tag);
          if (colNumber === undefined) {
            Browser.msgBox(`No column found for criterion '${criterion.name}'`);
            return;
          }

          rubricHeaderRangeValues[0][colNumber] = criterion.active
        });
      });

      rubricHeaderRange.setValues(rubricHeaderRangeValues);
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

      // 0 is the tag row, 1 is the human-readable header row
      headerRangeValues[0] = rosterHeaders.map(
        v => LibRubrics.GetSafeTagName(v)
      );
      headerRangeValues[1] = rosterHeaders;

      headerRange.setValues(headerRangeValues);

      gradingOverviewSheet.setColumnWidth(studentColumnSetup.colName, 150);
      gradingOverviewSheet.setColumnWidth(studentColumnSetup.colSurname, 150);
      gradingOverviewSheet.setColumnWidth(_ColClassroomID, 150);
      gradingOverviewSheet.setColumnWidth(_ColFullName, 150);
      gradingOverviewSheet.setColumnWidth(_ColOutput, 50);

      gradingOverviewSheet.hideColumns(studentColumnSetup.colUserId);
      gradingOverviewSheet.hideColumns(_ColCourseID);
      gradingOverviewSheet.hideColumns(studentColumnSetup.colEmail);
    }

    function GetTotalWidthNeeded(rubrics: LibRubrics.Rubric[]): number {
      return LibRubrics.CountCriteria(rubrics) + rubrics.length * 2;
    }

    function SetupRubricHeader(
      gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet,
      startColumn: number,
      rubrics: LibRubrics.Rubric[]) {
      // Get the range we need

      const totalWidthNeeded = GetTotalWidthNeeded(rubrics) + 4;

      const rubricHeaderRange = gradingOverviewSheet.getRange(
        1, startColumn,
        5, totalWidthNeeded
      );
      const rubricHeaderRangeValues = rubricHeaderRange.getValues();

      let currentCol = 0;

      // Go through the rubrics
      rubrics.forEach(rubric => {

        FormatRubricSingleHeader(gradingOverviewSheet, startColumn + currentCol, rubric);

        // Set rubric heading
        rubricHeaderRangeValues[_RowRubricTitle - 1][currentCol]
          = rubric.name;

        // Insert criteria info
        rubric.criteria.forEach(criteria => {
          rubricHeaderRangeValues[_RowHeading - 1][currentCol]
            = criteria.name;
          rubricHeaderRangeValues[_RowTag - 1][currentCol]
            = criteria.tag;
          rubricHeaderRangeValues[_RowGrade - 1][currentCol]
            = criteria.grade;
          rubricHeaderRangeValues[_RowCriteriaActive - 1][currentCol]
            = criteria.active;

          currentCol++;
        });

        // -- Setup Grade column for rubric
        rubricHeaderRangeValues[_RowTag - 1][currentCol]
          = rubric.gradeTag;
        rubricHeaderRangeValues[_RowHeading - 1][currentCol]
          = "Grade";
        rubricHeaderRangeValues[_RowCriteriaActive - 1][currentCol]
          = true;


        currentCol += 2;
      });

      // -- Comment header
      const commentColNum = currentCol;
      rubricHeaderRangeValues[_RowHeading - 1][currentCol] = "Comment";
      rubricHeaderRangeValues[_RowTag - 1][currentCol] = "comment";

      currentCol += 2;

      // -- Response doc header
      const responseDocColNum = currentCol
      rubricHeaderRangeValues[_RowHeading - 1][responseDocColNum] = "RESPONSE";
      rubricHeaderRangeValues[_RowTag - 1][responseDocColNum] = PageResponse._ResponseDocTag;

      // -- Insert values into range
      rubricHeaderRange.setValues(rubricHeaderRangeValues);

      // -- Set column widths
      gradingOverviewSheet.setColumnWidth(startColumn + commentColNum, 200);
      gradingOverviewSheet.setColumnWidth(startColumn + commentColNum + 1, 20);
      gradingOverviewSheet.setColumnWidth(startColumn + responseDocColNum, 200);
    }

    function FormatRubricSingleHeader(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet, rubricColumn: number, rubric: LibRubrics.Rubric) {
      // -- Rubric titles visuals
      let rubricTitleRange = gradingOverviewSheet.getRange(
        _RowRubricTitle, rubricColumn,
        1, rubric.criteria.length + 1);
      rubricTitleRange.merge();
      rubricTitleRange.setFontWeight("bold");

      // -- Criteria active checkboxes
      let criteriaCheckboxRange = gradingOverviewSheet.getRange(
        _RowCriteriaActive, rubricColumn,
        1, rubric.criteria.length + 1);
      criteriaCheckboxRange.insertCheckboxes();

      // -- Column widths
      gradingOverviewSheet.setColumnWidths(
        rubricColumn, rubric.criteria.length, 100
      );
      gradingOverviewSheet.setColumnWidth(
        rubricColumn + 1 + rubric.criteria.length, 20
      );
    }

    function FormatHeader(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet) {
      let lastCol = gradingOverviewSheet.getLastColumn();
      let headingRange = gradingOverviewSheet.getRange(
        _RowHeading, 1, 1,
        lastCol);
      let tagRange = gradingOverviewSheet.getRange(
        _RowTag, 1, 1,
        lastCol);

      headingRange.setFontWeight("bold");
      headingRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      tagRange.setFontSize(8)
        .setFontStyle("italic")
        .setWrap(true)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      tagRange.setFontStyle("italic");
    }
  }

  export function UpdateActiveCriteriaToTemplate() {
    // TODO: Implement
  }

  /* ---------------------------------------------------------------------------
    TRANSFERRING DATA
  ----------------------------------------------------------------------------*/
  //#region Transferring

  /**
   * Retrieve basic shallow info of all students from a grading overview sheet
   * Does not include data ranges or rubrics
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet whose overview to get students from
   * @returns {StudentData[]} an array of student data
   */
  export function GetAllStudentsData(
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ): LibStudents.StudentData[] {

    const studentsRange = gradingOverviewSheet.getRange(
      _RowDataStart, 1,
      gradingOverviewSheet.getLastRow() - _RowDataStart + 1,
      gradingOverviewSheet.getFrozenColumns()
    )

    return LibStudents.GetStudentsDataFromValues(studentsRange.getValues(), studentColumnSetup);
  }

  
  export function GetStudentDataRange(
    userID: string,
    colDataStart: number,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) : LibGSheets.RangeValuePair | undefined {

    // -- FIND THE RIGHT STUDENT
    const studentsData = GetAllStudentsData(gradingOverviewSheet);

    let studentRowNum = studentsData.findIndex(student => student.id === userID);
    if (studentRowNum < 0) { Browser.msgBox("Student ID not found"); return undefined; };

    // -- Get the student's data
    const studentData = GetGradingDataRow(
      studentRowNum, colDataStart, gradingOverviewSheet
    );

    return studentData;
  }

  /**
   * Insert rubric data for a specified user in the grading overview sheet
   * @param userID {string}
   * @param data {PageStudentDetails.GradingData}
   * @param gradingOverviewSheet {GoogleAppsScript.Spreadsheet.Sheet}
   * @returns 
   */
  export function InsertRubricData(
    userID: string, // TODO: Extract get-student-range; replace this w/ target range?
    data: LibStudents.GradingData,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {

    // -- PREPARE

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

    // Find the first column that contains a criteria
    //  Used b/ we don't want to destroy formulas of preeceding columns
    //  TODO: Examine how we can avoid this

    const colDataStart = GetFirstCriteriaColumnIndex(tagColNumbers, data) + 1;

    // -- FIND THE RIGHT STUDENT
    const studentData = GetStudentDataRange(userID, colDataStart, gradingOverviewSheet);
    if (studentData === undefined) return;

    // -- INSERT DATA

    // -- Check if there are already values
    const numValues = studentData.values[0].filter(v => v.length != 0).length;
    if (numValues > 0) {
      const answer = Browser.msgBox(
        "Values already exist for that student. Overwrite?",
        Browser.Buttons.YES_NO
      );
      if (answer == "no") {
        return;
      }
    }


    // -- Go through all rubrics, insert checkmarks & grades
    data.rubrics.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // Find the column with a matching tag (including data start offset)
        let colNumber = (tagColNumbers.get(criterion.tag) ?? 0) - (colDataStart - 1);
        if (colNumber < 0) {
          Browser.msgBox(`No column found for criterion '${criterion.name}'`);
          return;
        }

        studentData.values[0][colNumber] = criterion.studentPassed ? "✔" : "✘";
      });

      // -- Set rubric grade
      let colNumber = (tagColNumbers.get(rubric.gradeTag) ?? 0) - (colDataStart - 1);
      if (colNumber < 0) return;

      studentData.values[0][colNumber] = rubric.studentGrade;
    });

    // -- Set comment

    // Find the right column, including data start offset
    let colNumber = (tagColNumbers.get("comment") ?? 0) - (colDataStart - 1); // TODO: This tag is bad b/c someone might use it accidentally
    if (colNumber < 0) {
      Browser.msgBox("No column found for comment");
    }
    else {
      // colNumber -= (colDataStart - 1);
      studentData.values[0][colNumber] = data.comment;
    }

    // -- Re-insert values
    studentData.range.setValues(studentData.values);
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
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet): LibStudents.StudentData | null {

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

    // -- Get the student data values

    // Find the right student
    const studentsData = GetAllStudentsData(gradingOverviewSheet);

    let studentRowNum = studentsData.findIndex(student => student.id === userID);
    if (studentRowNum < 0) { Browser.msgBox("Student ID not found"); return null; };

    const student = studentsData[studentRowNum];

    // Get the student's data
    const studentDataValues = gradingOverviewSheet.getRange(
      _RowDataStart + studentRowNum, // Student's row
      1,
      1, // only one row
      gradingOverviewSheet.getMaxColumns()
    ).getValues()[0].map(v => String(v));

    // -- Get the rubrics from the rubrics sheet
    student.gradingData = {
      rubrics: PageRubrics.GetRubrics(rubricsSheet),
      comment: ""
    }

    LibStudents.InsertRowDataIntoStudent(student, tagColNumbers, studentDataValues);

    return student;
  }

  //#endregion

  /* -----------------------------------------------------------------------------
    HELPER FUNCTIONS
  ------------------------------------------------------------------------------*/
  //#region helper functions

  function GetGradingDataRow(
    offset: number,
    colDataStart: number,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ): LibGSheets.RangeValuePair {

    const data = gradingOverviewSheet.getRange(
      _RowDataStart + offset,
      colDataStart,
      1, // only one row
      gradingOverviewSheet.getMaxColumns() - (colDataStart - 1)
    );

    return {
      range: data,
      values: data.getValues()
    }
  }

  /**
   * Creates a Map<string, number> from a grading overview sheet, where the keys
   *   are the tags and the values are the corresponding column number, counted 
   *   from colDataStart
   * @param gradingOverviewSheet {GoogleAppsScript.Spreadsheet.Sheet} The grading overview sheet
   * @param colDataStart {number} The column where grading/rubric/criteria data starts
   * @returns {Map<string, number>} The finished map
   */
  export function MakeTagColNumberMap(
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ): Map<string, number> {

    // Get the tags-row from the overview sheet
    const tagsValues = gradingOverviewSheet.getRange(
      _RowTag,
      1,
      1, // only one row
      gradingOverviewSheet.getMaxColumns()
    ).getValues();

    const tagColNumbers = new Map<string, number>();

    tagsValues[0].forEach((col, colNum) => {
      if (col) {
        tagColNumbers.set("" + col, colNum);
      }
    });

    return tagColNumbers;
  }


  export function GetFirstCriteriaColumnIndex(
    tagColNumbers: Map<string, number>,
    gradingData: LibStudents.GradingData
  ): number {
    const allCriteria = gradingData.rubrics.flatMap((rubric) => rubric.criteria);
    return Math.min(...allCriteria.map(criteria => tagColNumbers.get(criteria.tag) ?? 0));
  }

  //#endregion
}