import { LibRubrics } from "../libs/rubrics";
import { LibGSheets } from "../libs/sheets"
import { PageGradingOverview } from "./gradingoverview";
import { PageStudentDetails } from "./studentdetails";

export namespace PageStudentGrading {

  const _StudentGradingSheetName = "STUDENTGRADE";

  const _ColRubric: number = 1;
  const _ColCriteria: number = 2;
  const _ColColnum: number = 3;
  const _ColCheckmark: number = 4;
  const _ColGrade: number = 5;
  const _ColActive: number = 6;

  const _RowHeader: number = 3;
  const _EditBoxColor: number[] = [217, 234, 211];

  /**
   * Add a student grading sheet to a spreadsheet, based on its grading overview data (students, rubrics)
   * @param spreadsheet The spreadsheet to add the student grading sheet to
   * @returns 
   */
  export function Setup(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
  ) {

    // -- CONFIG
    const setup: PageStudentDetails.SheetSetup = {
      ColRubric: 1,
      ColCriteria: 2,
      ColColnum: 3,
      ColCheckmark: 4,
      ColGrade: 5,
      ColActive: 6,
      RowHeader: 3,

      IncludeCheckboxCol: true,
      IncludeGradeCol: false,
      IncludeGradeLine: true,
      IncludeCommentLine: true,

      CheckboxColType: "checkable",
      CheckboxColColorized: true,
    }

    // -- PREP
    const studentGradingSheet = LibGSheets.CreateOrGetSheet(
      _StudentGradingSheetName,
      spreadsheet, true
    )

    const gradingOverviewSheet = PageGradingOverview.GetGradingOverviewSheet(spreadsheet);

    if (!studentGradingSheet || !gradingOverviewSheet) return;


    // -- GET DATA
    const rubrics = PageGradingOverview.GetRubrics(gradingOverviewSheet);
    const students = PageGradingOverview.GetStudentsData(gradingOverviewSheet);

    // -- CLEAR & SET SIZE
    LibGSheets.ClearSheet(studentGradingSheet);

    const totalHeight = _RowHeader
      + LibRubrics.CountCriteria(rubrics)
      + rubrics.length * 2 // Space for grade + spacing
      + 3; // Space for comment block

    LibGSheets.SetSheetSize(studentGradingSheet, 8, totalHeight);

    // -- SETUP BLOCKS

    PageStudentDetails.SetupHeaderBlock(studentGradingSheet, students, setup);
    SetupHelpers.SetupRubricsBlock(studentGradingSheet, rubrics);

    // -- SET WIDTHS
    studentGradingSheet
      .setColumnWidth(_ColRubric, 223)
      .setColumnWidth(_ColCriteria, 275)
      .setColumnWidth(_ColGrade, 70)
      .setColumnWidth(_ColActive, 70)
      .hideColumns(_ColColnum);
  }

  namespace SetupHelpers {

    /**
     * Add a block of rubrics & criteria
     * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The sheet to add rubrics block to
     * @param {LibRubrics.Rubric[]} rubrics - The rubrics to add rows etc for
     */
    export function SetupRubricsBlock(
      studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
      rubrics: LibRubrics.Rubric[]
    ) {

      let rubricStartRow = _RowHeader + 1;

      const dataRange = studentGradingSheet.getRange(rubricStartRow, 1, studentGradingSheet.getMaxRows() - _RowHeader, 6);
      const dataValues = dataRange.getValues();

      // -- RUBRICS ROWS
      let row = 0;

      rubrics.forEach(rubric => {
        let rubricBlockStartRow = rubricStartRow + row;

        dataValues[row][0] = rubric.name;

        // Insert rows from criteria
        rubric.criteria.forEach(criteria => {
          dataValues[row][_ColCriteria - 1] = criteria.name;
          dataValues[row][_ColColnum - 1] = criteria.columnNumber;
          dataValues[row][_ColCheckmark - 1] = "✘";
          dataValues[row][_ColGrade - 1] = criteria.grade;
          dataValues[row][_ColActive - 1] = criteria.active;
          row++;
        });

        // "Grade" on its own row
        dataValues[row][_ColCriteria - 1] = "Grade";
        dataValues[row][_ColColnum - 1] = rubric.criteria.slice(-1)[0].columnNumber + 1;
        dataValues[row][_ColActive - 1] = true;

        row += 2;

        // When done, format the block
        FormatRubricBlock(studentGradingSheet, rubric.criteria.length, rubricBlockStartRow)
      });

      // -- COMMENT ROW
      dataValues[row + 1][_ColCriteria - 1] = "Comment";

      // Offset is 3 because last criteria's colnr + last grade colnr + 2.
      const commentColNr = 3 + (rubrics.at(-1)?.criteria.at(-1)?.columnNumber ?? 0);
      dataValues[row + 1][_ColColnum - 1] = commentColNr.toString();

      dataRange.offset(row + 1, _ColCriteria - 1, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold")
        .offset(0, 2, 1, 3) // get writing box
        .setBackgroundRGB(_EditBoxColor[0], _EditBoxColor[1], _EditBoxColor[2])
        .merge();

      // -- FINALIZING
      dataRange.setValues(dataValues);

      // General formatting
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment("top");

      // ADD FILTER
      SetFilter(rubrics, dataRange);
    }

    /**
     * Add formatting to a rubric's block
     * @param studentGradingSheet - The sheet where the formattin's taking place
     * @param numCriteria - Number of criteria rows
     * @param rubricBlockStartRow - The row where the rubric's block starts
     */
    function FormatRubricBlock(
      studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
      numCriteria: number,
      rubricBlockStartRow: number
    ) {

      // Rubric label block
      studentGradingSheet.getRange(rubricBlockStartRow, _ColRubric, numCriteria + 1, 1)
        .merge()
        .setBackground("#EFEFEF")
        .setFontWeight("bold");

      // Checkboxes
      studentGradingSheet.getRange(rubricBlockStartRow, _ColCheckmark, numCriteria, 1)
        .setHorizontalAlignment("center")
        .insertCheckboxes("✔", "✘");

      // Grade sub-block
      studentGradingSheet.getRange(rubricBlockStartRow + numCriteria, _ColCriteria, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold");

      studentGradingSheet.getRange(rubricBlockStartRow + numCriteria, _ColCheckmark, 1, 1)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackgroundRGB(_EditBoxColor[0], _EditBoxColor[1], _EditBoxColor[2])
    }


    /**
     * Add a filter to a rubrics block
     * @param rubrics - The rubrics data
     * @param dataRange - The range where the rubric blocks were added
     */
    function SetFilter(rubrics: LibRubrics.Rubric[], dataRange: GoogleAppsScript.Spreadsheet.Range) {
      // Count number of criteria
      const totalHeight = LibRubrics.CountCriteria(rubrics)
        + rubrics.length * 2; // Add 1 for the grade and 1 for spacing, for each rubric

      let filterRange = dataRange.offset(-1, 0, totalHeight);
      let filter = filterRange.createFilter();
      const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["FALSE"]);
      filter.setColumnFilterCriteria(_ColActive, criteria);
    }
  }


  export function GetStudentGradingSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_StudentGradingSheetName);
  }

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
      studentGradingSheet.getMaxRows() - _RowHeader + 1 // TODO: How to improve speed?
    )

    const checkmarkValues = checkmarkRange.getValues().map(row => {
      if (row[0] === "✔" || row[0] === "✘") return ["✘"]
      else return [""];
    });

    checkmarkRange.setValues(checkmarkValues);

    ClearSelectedUserId(studentGradingSheet);
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