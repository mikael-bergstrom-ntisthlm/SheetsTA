import { LibGClassroom } from "../libs/classroom.js";
import { LibConfig } from "../libs/config.js";
import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js";
import { PageRubrics } from "./rubrics.js";
import { PageStudentDetails } from "./studentdetails.js";

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
      // TODO: Add some sort of warning if there's already data

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

      fullNameRange.setFormula(`=${String.fromCharCode(64 + _ColSurname)}${_RowDataStart} & " " & ${String.fromCharCode(64 + _ColName)}${_RowDataStart}`);

      fullNameRange.setBackground("#d9d9d9");
    }

    export function UpdateActiveCriteriaFromTemplate(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
      const gradingOverviewSheet = GetDefaultGradingOverviewSheet(spreadsheet);
      const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);

      if (!gradingOverviewSheet || !rubricsSheet) return;

      let rubrics = PageRubrics.GetRubrics(rubricsSheet);

      let startColumn = gradingOverviewSheet.getFrozenColumns();
      const totalWidthNeeded = GetTotalWidthNeeded(rubrics) + 4;

      // Get the range we need
      let rubricHeaderRange = gradingOverviewSheet.getRange(
        _RowCriteriaActive, startColumn,
        1, startColumn + totalWidthNeeded
      );
      let rubricHeaderRangeValues = rubricHeaderRange.getValues();

      // -- Make a map of which column belongs to which tag
      const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet, startColumn);

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

      // TODO: Use the consts for col-numbers (wtf did I mean by this?)
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
      rubricHeaderRangeValues[_RowTag - 1][responseDocColNum] = "responsedoc";

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
  export function GetStudentsData(gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const studentValues = gradingOverviewSheet.getRange(
      _RowDataStart, 1,
      gradingOverviewSheet.getLastRow() - _RowDataStart + 1,
      gradingOverviewSheet.getFrozenColumns()
    ).getValues()

    const studentsData: PageStudentDetails.StudentData[] = [];

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
   * Insert rubric data for a specified user in the grading overview sheet
   * @param userID {string}
   * @param data {PageStudentDetails.GradingData}
   * @param gradingOverviewSheet {GoogleAppsScript.Spreadsheet.Sheet}
   * @returns 
   */
  export function InsertRubricData(
    userID: string,
    data: PageStudentDetails.GradingData,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {

    // TODO: Make this more precise
    const colDataStart = gradingOverviewSheet.getFrozenColumns() + 1;

    // -- Find the right student
    const studentsData = GetStudentsData(gradingOverviewSheet);

    let studentRowNum = studentsData.findIndex(student => student.id === userID);
    if (studentRowNum < 0) { Browser.msgBox("Student ID not found"); return null; };

    // -- Get the student's data
    const studentData = GetGradingDataRow(
      studentRowNum, colDataStart, gradingOverviewSheet
    );

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

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet, colDataStart);

    // -- Go through all rubrics, insert checkmarks & grades
    data.rubrics.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // Find the column with a matching tag
        const colNumber = tagColNumbers.get(criterion.tag);
        if (colNumber === undefined) {
          Browser.msgBox(`No column found for criterion '${criterion.name}'`);
          return;
        }

        studentData.values[0][colNumber] = criterion.studentPassed ? "✔" : "✘";
      });

      // -- Set rubric grade
      const colNumber = tagColNumbers.get(rubric.gradeTag);
      if (colNumber === undefined) return;
      studentData.values[0][colNumber] = rubric.studentGrade;
    });

    const colNumber = tagColNumbers.get("comment"); // TODO: This tag is bad b/c someone might use it accidentally
    if (colNumber === undefined) {
      Browser.msgBox("No column found for comment");
    }
    else {
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
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet): PageStudentDetails.StudentData | null {

    // TODO: Make this more precise
    const colDataStart = gradingOverviewSheet.getFrozenColumns() + 1;

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet, colDataStart);

    // -- Get the student data values

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

    // -- Get the rubrics from the rubrics sheet
    student.gradingData = {
      rubrics: PageRubrics.GetRubrics(rubricsSheet),
      comment: "" // FIXME: Get the actual comment
    }

    // Get the tags-row from the overview sheet
    const tagsValues = gradingOverviewSheet.getRange(
      _RowTag, colDataStart,
      1, gradingOverviewSheet.getMaxColumns() - colDataStart
    ).getValues();

    // Go through the rubrics
    student.gradingData.rubrics.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // Find the column with a matching tag
        const colNumber = tagColNumbers.get(criterion.tag);
        if (colNumber === undefined) {
          Browser.msgBox(`No column found for criterion '${criterion.name}'`);
          return;
        }

        criterion.studentPassed =
          studentDataValues[0][colNumber] == "✔";
      });

      // Find column of rubric's overall grade
      const gradeCol = tagColNumbers.get(rubric.gradeTag)
      if (gradeCol === undefined) {
        Browser.msgBox(`No grade column found for rubric "${rubric.name}"`);
      } else {
        // Save rubric grade
        rubric.studentGrade = studentDataValues[0][gradeCol];
      }
    });

    // Get comment
    const colNumber = tagColNumbers.get("comment");
    if (colNumber === undefined) {
      Browser.msgBox("No column found for comment");
    }
    else {
      student.gradingData.comment = studentDataValues[0][colNumber];
    }

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
      gradingOverviewSheet.getMaxColumns() - colDataStart
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
  function MakeTagColNumberMap(
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet,
    colDataStart: number
  ): Map<string, number> {

    // Get the tags-row from the overview sheet
    const tagsValues = gradingOverviewSheet.getRange(
      _RowTag,
      colDataStart,
      1, // only one row
      gradingOverviewSheet.getMaxColumns() - colDataStart
    ).getValues();

    const tagColNumbers = new Map<string, number>();

    tagsValues[0].forEach((col, colNum) => {
      tagColNumbers.set("" + col, colNum);
    });

    return tagColNumbers;
  }

  //#endregion
}