import { LibGSheets } from "../libs/sheets.js";
import { PageRubrics } from "./rubrics.js";
import { PageStudentDetails } from "./studentdetails.js";


export namespace PageResponse {

  const _ResponseTemplateSheetName = "_TEMPLATERESPONSE";

  // -- CONFIG
  export const setup: PageStudentDetails.SheetSetup = {
    ColRubric: 1,
    ColCriteria: 2,
    ColTag: 3,
    ColCheckmark: 4,
    ColGrade: -1,
    ColActive: 5,
    RowHeaderHeight: 4,
    RowHeaderName: 1,
    RowHeaderComment: 2,

    CommentFooter: false,
    GradeForEachRubric: false,
  }

  // TODO: CURRENT PROJECT
  export function Setup(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  ) {



    // -- PREP
    const responseTemplate = LibGSheets.CreateOrGetSheet(
      _ResponseTemplateSheetName,
      spreadsheet, true
    )

    const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);

    if (!responseTemplate || !rubricsSheet) {
      SpreadsheetApp.getUi().alert("At least one sheet not found (response template, rubrics)");
      return;
    }

    LibGSheets.ClearSheet(responseTemplate);

    // -- GET DATA

    const rubrics = PageRubrics.GetRubrics(rubricsSheet);

    PageStudentDetails.SetupHeaderBlock(
      responseTemplate, [],
      setup
    )

    PageStudentDetails.SetupRubricsBlock(
      responseTemplate,
      rubrics,
      setup
    )

    // -- HIDE TAG COLUMN
    if (setup.ColTag > 0) {
      responseTemplate.hideColumns(setup.ColTag);
    }

    // -- SET WIDTHS
    if (setup.ColRubric > 0) {
      responseTemplate
        .setColumnWidth(setup.ColRubric, 223)
    }
    if (setup.ColCriteria > 0) {
      responseTemplate
        .setColumnWidth(setup.ColCriteria, 275);
    }
  }

  export function GenerateResponseDocument(
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet,
    userId: string
  ) {

    // Test: generate new Sheets document and add some stuff to it
    //       Add url to correct column


    let overviewSpreadsheetId = gradingOverviewSheet.getParent().getId();
    let folder = DriveApp.getFileById(overviewSpreadsheetId).getParents().next();

    let newSheet = SpreadsheetApp.create("test");
    let newFile = DriveApp.getFileById(newSheet.getId());
    newFile.moveTo(folder);

  }
}