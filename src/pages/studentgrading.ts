import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js"
import { PageGradingOverview } from "./gradingoverview.js";
import { PageRubrics } from "./rubrics.js";
import { PageStudentDetails } from "./studentdetails.js";

export namespace PageStudentGrading {

  const _StudentGradingSheetName = "STUDENTGRADE";

  // -- CONFIG
  export const setup: PageStudentDetails.SheetSetup = {
    ColRubric: 1,
    ColCriteria: 2,
    ColTag: 3,
    ColCheckmark: 4,
    ColGrade: 5,
    ColActive: 6,
    ColHeaderData: 2,
    RowHeaderHeight: 3,
    RowHeaderName: 1,
    RowHeaderComment: -1,

    CommentFooter: true,
    GradeForEachRubric: true
  }


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
        .setColumnWidth(setup.ColRubric, 223)
        .setColumnWidth(setup.ColCriteria, 275)
        .setColumnWidth(setup.ColGrade, 70)
        .setColumnWidth(setup.ColActive, 70)
      studentGradingSheet
        .hideColumns(setup.ColTag);

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

    const nameCellValue: string = studentGradingSheet.getRange(setup.RowHeaderName, setup.ColHeaderData).getValue();

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
    CLEARING & RESETTING
  ----------------------------------------------------------------------------*/
  //#region Clearing and resetting

  /**
   * Clear a student grading sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The student grading sheet
   */
  export function ClearGrading(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const checkmarkRange = studentGradingSheet.getRange(
      setup.RowHeaderHeight + 1,
      setup.ColCheckmark,
      studentGradingSheet.getMaxRows() - setup.RowHeaderHeight + 1
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
