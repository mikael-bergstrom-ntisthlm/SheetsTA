import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js"
import { LibStudents } from "../libs/students.js";
import { PageGradingOverview } from "./gradingoverview.js";
import { PageRubrics } from "./rubrics.js";
import { PageStudentDetails } from "./studentdetails.js";

export namespace PageStudentGrading {

  const _StudentGradingSheetName = "STUDENTGRADE";

  const _ColName: number = 2;
  const _RowName: number = 1;

  const _ColRubric: number = 1;
  const _ColCriteria: number = 2;
  const _ColTag: number = 3;
  const _ColCheckmark: number = 4;
  const _ColGrade: number = 5;
  const _ColActive: number = 6;

  const _RowHeader: number = 3;


  export namespace Setup {
    /**
     * Add a student grading sheet to a spreadsheet, based on its grading overview data (students, rubrics)
     * @param spreadsheet The spreadsheet to add the student grading sheet to
     * @returns 
     */
    export function Setup(
      spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
    ) {

      // TODO: Decide wtf to do here – global config or config object?
      // -- CONFIG
      const setup: PageStudentDetails.SheetSetup = {
        ColRubric: _ColRubric,
        ColCriteria: _ColCriteria,
        ColTag: _ColTag,
        ColCheckmark: _ColCheckmark,
        ColGrade: _ColGrade,
        ColActive: _ColActive,
        RowHeaderHeight: _RowHeader,
        RowHeaderName: 1,
        RowHeaderComment: -1,

        CommentFooter: true,
        GradeForEachRubric: true
      }

      // -- PREP
      const studentGradingSheet = LibGSheets.CreateOrGetSheet(
        _StudentGradingSheetName,
        spreadsheet, true
      )

      const gradingOverviewSheet = PageGradingOverview.GetDefaultGradingOverviewSheet(spreadsheet);
      const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);

      if (!studentGradingSheet || !gradingOverviewSheet || !rubricsSheet) {
        SpreadsheetApp.getUi().alert("At least one sheet not found (student grading, overview, rubrics)");
        return;
      }

      // -- GET DATA

      const rubrics = PageRubrics.GetRubrics(rubricsSheet);
      const students = PageGradingOverview.GetAllStudentsData(gradingOverviewSheet);

      // -- CLEAR & SET SIZE
      LibGSheets.ClearSheet(studentGradingSheet);

      const totalHeight = setup.RowHeaderHeight
        + LibRubrics.CountCriteria(rubrics)
        + rubrics.length * 2 // Space for grade + spacing
        + 3; // Space for comment block

      const totalWidth = PageStudentDetails.GetHighestColumnNumber(setup) + 1;

      LibGSheets.SetSheetSize(studentGradingSheet, totalWidth, totalHeight);

      // -- SETUP BLOCKS
      PageStudentDetails.SetupHeaderBlock(studentGradingSheet, students, setup);
      PageStudentDetails.SetupRubricsBlock(studentGradingSheet, rubrics, setup);

      // -- SET WIDTHS
      studentGradingSheet
        .setColumnWidth(_ColRubric, 223)
        .setColumnWidth(_ColCriteria, 275)
        .setColumnWidth(_ColGrade, 70)
        .setColumnWidth(_ColActive, 70)
      studentGradingSheet
        .hideColumns(_ColTag);

    }

  }

  export function GetDefaultStudentGradingSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_StudentGradingSheetName);
  }

  /**
   * Get the ID of the currently selected user of a student grading sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The sheet
   * @returns {string} An ID
   */
  export function GetSelectedUserId(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): string {

    const nameCellValue: string = studentGradingSheet.getRange(_RowName, _ColName).getValue();

    if (nameCellValue == "") {
      SpreadsheetApp.getUi().alert("No selection!");
      return "";
    }

    let pair = nameCellValue.split("|");
    if (pair.length != 2 || pair[1] === "") {
      SpreadsheetApp.getUi().alert("Invalid selection!");
      return "";
    }

    return pair[1].trim();
  }

  /* ---------------------------------------------------------------------------
    TRANSFERRING DATA
  ----------------------------------------------------------------------------*/
  //#region Transferring

  /**
   * Insert Student data from some other source, using criteria tags to match
   * with student grading sheet rows
   * @param student The student data to insert
   * @param studentGradingSheet The sheet to insert it into
   */
  export function InsertStudentDataRubrics(
    student: LibStudents.StudentData,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet
  ): void {

    if (!student.gradingData) {
      Browser.msgBox("Student has no data!");
      return;
    }

    const localData = GetRubricsData(studentGradingSheet);

    // Setup quick index of tags and row numbers for easy lookup
    const tagRowNumbers = new Map<string, number>();

    localData.values.forEach((row, rowNum) => {
      tagRowNumbers.set("" + row[_ColTag - 1], rowNum);
    });

    // Go through the rubrics, get grades from local data
    student.gradingData.rubrics.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // Find the row with the corresponding tag
        const rowNum = tagRowNumbers.get(criterion.tag);
        if (rowNum === undefined) {
          Browser.msgBox(`No row found for criterion '${criterion.name}'`);
          return;
        }

        // Set the row's checkmark status
        localData.values[rowNum][_ColCheckmark - 1] =
          criterion.studentPassed ? "✔" : "✘";
      });

      // Set the grade
      const rowNum = tagRowNumbers.get(rubric.gradeTag);
      if (rowNum === undefined) return;
      localData.values[rowNum][_ColCheckmark - 1] = rubric.studentGrade;
    });

    // Set the comment
    const rowNum = tagRowNumbers.get("comment");
    if (rowNum) {
      localData.values[rowNum][_ColCheckmark - 1] = student.gradingData.comment;
    }

    
    // Insert the data
    localData.range.setValues(
      localData.values
    )
  }


  export function GetStudentGradingData(
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet
  ): LibStudents.GradingData {

    // Get rubrics from rubrics page
    const data: LibStudents.GradingData = {
      rubrics: PageRubrics.GetRubrics(rubricsSheet),
      comment: ""
    }

    if (data.rubrics.length == 0) { Browser.msgBox("No rubrics found") }

    // Get the local values
    const localData = GetRubricsData(studentGradingSheet);

    // Setup quick index of tags and row numbers for easy lookup
    const tagRowNumbers = new Map<string, number>();

    localData.values.forEach((row, rowNum) => {
      tagRowNumbers.set("" + row[_ColTag - 1], rowNum);
    });

    // Go through the rubrics, set grades from local data
    data.rubrics.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // Find the row with the corresponding tag
        const rowNum = tagRowNumbers.get(criterion.tag);
        if (rowNum === undefined) {
          Browser.msgBox(`No row found for criterion '${criterion.name}'`);
          return
        };

        // Set passed/not passed
        criterion.studentPassed =
          localData.values[rowNum][_ColCheckmark - 1] == "✔" ? true : false;
      });

      // Set the grade
      const rowNum = tagRowNumbers.get(rubric.gradeTag);
      if (rowNum === undefined) return;
      rubric.studentGrade = localData.values[rowNum][_ColCheckmark - 1];
    });

    // -- Get the comment
    const rowNum = tagRowNumbers.get("comment");
    if (rowNum) {
      data.comment = localData.values[rowNum][_ColCheckmark - 1];
    }

    // Return the data
    return data;
  }

  /**
   * Get the entire rubrics block (range+values) of a student grading sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The student grading sheet
   * @returns {RangeValuePair} A value-range pair
   */
  function GetRubricsData(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): LibGSheets.RangeValuePair {

    const gradingDataRange = studentGradingSheet
      .getRange(_RowHeader + 1, 1, // Start at the row below the header
        studentGradingSheet.getLastRow() - _RowHeader, // Get all the rows, minus the header
        Math.max(_ColActive, _ColCheckmark, _ColCriteria, _ColGrade, _ColRubric)); // Find the rightmost column

    return {
      values: gradingDataRange.getValues(),
      range: gradingDataRange
    };
  }

  //#endregion

  /* ---------------------------------------------------------------------------
    CLEARING & RESETTING
  ----------------------------------------------------------------------------*/
  //#region Clearing and resetting

  /**
   * Clear a student grading sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The student grading sheet
   */
  export function ClearGrading(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const checkmarkRange = studentGradingSheet.getRange(
      _RowHeader + 1,
      _ColCheckmark,
      studentGradingSheet.getMaxRows() - _RowHeader + 1
    )

    const checkmarkValues = checkmarkRange.getValues().map(row => {
      return [GetClearGradingFor(row[0])]
    });

    checkmarkRange.setValues(checkmarkValues);
    ClearSelectedUserId(studentGradingSheet);
  }

  function GetClearGradingFor(currentValue: string) {
    return (currentValue === "✔" || currentValue === "✘")
      ? "✘"
      : ""
  }

  /**
   * Clear student selection of a student grading sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The student grading sheet
   */
  function ClearSelectedUserId(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    studentGradingSheet.getRange(1, 2).setValue("");
  }

  //#endregion


}
