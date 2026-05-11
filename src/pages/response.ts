import { LibGSheets } from "../libs/sheets.js";
import { LibStudents } from "../libs/students.js";
import { PageRubrics } from "./rubrics.js";
import { PageStudentDetails } from "./studentdetails.js";

export namespace PageResponse {

  const _ResponseTemplateSheetName = "_TEMPLATERESPONSE";
  const _ResponseSheetDetailsName = "DETAILS";

  export const _ResponseDocTag = "responsedoc";

  // -- CONFIG
  export const setup: PageStudentDetails.SheetSetup = {
    ColRubric: 1,
    ColCriteria: 2,
    ColTag: 3,
    ColCheckmark: 4,
    ColGrade: -1, // TODO: Make this *optional*
    ColActive: 5,
    ColHeaderData: 2,
    RowHeaderHeight: 4,
    RowHeaderName: 1,
    RowHeaderComment: 2,

    CommentFooter: false,
    GradeForEachRubric: false,
  }

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

    // TODO: Crop unnecessary rows and columns

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

  export function GetDefaultResponseTemplateSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_ResponseTemplateSheetName);
  }


  export function GetOrCreateDetailsSheet(
    studentResponseSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    responseTemplateSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {
    let detailsSheet = studentResponseSpreadsheet.getSheetByName(_ResponseSheetDetailsName);
    if (detailsSheet === null) {
      detailsSheet = responseTemplateSheet.copyTo(studentResponseSpreadsheet);
      detailsSheet.setName(_ResponseSheetDetailsName);
    }
    studentResponseSpreadsheet.setActiveSheet(detailsSheet);
    studentResponseSpreadsheet.moveActiveSheet(1);

    return detailsSheet;
  }


  /* -----------------------------------------------------------------------------
  RESPONSE DOCUMENT GENERATION
------------------------------------------------------------------------------*/
  //#region response doc gen

  // TODO: CURRENT PROJECT
  /**
   * 
   * @param rowBlocks An array of Ranges; expected to already be full-width
   * @param tagColNumbers A map of tags and columns matching the rowBlocks
   * @param targetFolder The folder where response documents are moved to
   * @param responseTemplateSheet The sheet to use as template
   * @param rubricsSheet The sheet containing rubric data
   * @returns 
   */
  export function GenerateResponseDocuments(
    rowBlocks: GoogleAppsScript.Spreadsheet.Range[],
    tagColNumbers: Map<string, number>,
    studentColumnSetup: LibStudents.StudentColumnSetup,
    targetFolder: GoogleAppsScript.Drive.Folder,
    responseTemplateSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {

    // -- PREPARE

    // -- Get parent spreadsheet; for toasts
    const spreadsheet = responseTemplateSheet.getParent();

    // -- Make a map of which column belongs to which tag
    // const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

    // -- Make sure there's a column for the response doc URL
    const responseColNum = tagColNumbers.get(_ResponseDocTag);
    if (responseColNum === undefined) {
      Browser.msgBox(`No response document column found! \\nNeeds to have the tag ${_ResponseDocTag}`);
      return;
    }

    const rubrics = PageRubrics.GetRubrics(rubricsSheet);

    // -- PROCESS

    rowBlocks.forEach(rowBlock => {
      const rowBlockValues = rowBlock.getValues();
      const students = LibStudents.GetStudentsDataFromValues(rowBlockValues, studentColumnSetup);

      // -- Go through each student of the current block
      for (let i = 0; i < students.length; i++) {

        // -- PREP STUDENT INCLUDING RUBRICS
        const student = students[i];

        student.gradingData = {
          comment: "",
          rubrics: JSON.parse(JSON.stringify(rubrics))
        }

        LibStudents.InsertRowDataIntoStudent(student, tagColNumbers, rowBlockValues[i]);

        // -- PREP DOCUMENT
        let responseDocUrl: string = rowBlockValues[i][responseColNum];

        let studentResponseSpreadsheet =
          GetOrCreateStudentResponseSpreadsheet(student, responseDocUrl, targetFolder);

        if (studentResponseSpreadsheet === undefined) return;

        // Get the right sheet, if it exists
        let responseSheet = PageResponse.GetOrCreateDetailsSheet(
          studentResponseSpreadsheet,
          responseTemplateSheet
        );

        // -- COMBINE STUDENT DATA WITH RESPONSE DOC
        PageStudentDetails.InsertStudentDataRubrics(student, responseSheet, PageResponse.setup);

        // Name field
        responseSheet.getRange(PageResponse.setup.RowHeaderName, PageResponse.setup.ColHeaderData)
          .setValue(`${student.name} ${student.surname}`);

        // Comment field
        responseSheet.getRange(PageResponse.setup.RowHeaderComment, PageResponse.setup.ColHeaderData)
          .setValue(student.gradingData.comment);


        // -- INSERT NEW URL
        const newUrl = studentResponseSpreadsheet.getUrl();
        if (newUrl !== responseDocUrl) {
          let responseBlock = rowBlock.offset(
            i,
            responseColNum,
            1, 1
          );

          responseBlock.setValue(newUrl)
        }

        spreadsheet.toast(`Document for ${student.name} ${student.surname} updated`);
      }

    });
  }

  /**
   * Get or create a spreadsheet file for the given student. Checks if the given responseDocUrl
   * leads to an existing file; if it does returns a reference to it.
   * 
   * If it doesn't lead to an existing valid Drive file, a new file is created using the 
   * student data to give the new file its file name. A reference to the new file is returned.
   * 
   * The function does not handle the contents of the file at all.
   * @param student The student for whom to create the document
   * @param responseDocUrl The existing potential document URL
   * @param targetFolder The folder where the new file is created (or the found file moved)
   * @returns The existing or the new file (or 'undefined' if no file was created or found)
   */
  function GetOrCreateStudentResponseSpreadsheet(
    student: LibStudents.StudentData,
    responseDocUrl: string,
    targetFolder: GoogleAppsScript.Drive.Folder
  ): GoogleAppsScript.Spreadsheet.Spreadsheet | undefined {

    const responseSpreadsheetName = `Response ${student.surname} ${student.name}`;

    let studentResponseSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | undefined = undefined;
    let studentResponseSpreadsheetFile: GoogleAppsScript.Drive.File | undefined = undefined;

    if (responseDocUrl !== "") {
      try {
        studentResponseSpreadsheet = SpreadsheetApp.openByUrl(responseDocUrl);
        studentResponseSpreadsheetFile = DriveApp.getFileById(studentResponseSpreadsheet.getId());

        // Disregard if trashed
        if (studentResponseSpreadsheetFile.isTrashed()) {
          studentResponseSpreadsheet = undefined;
          studentResponseSpreadsheetFile = undefined;
        }
      }
      catch {
        const overwrite = Browser.msgBox(`Student "${student.name} ${student.surname} has something in the response doc column, but it doesn't seem to be the url of a Spreadsheet document\\nDo you want to overwrite this content?"`, Browser.Buttons.YES_NO);
        if (overwrite === "no") return undefined;
      }
    }

    if (studentResponseSpreadsheet === undefined || studentResponseSpreadsheetFile === undefined) {
      studentResponseSpreadsheet = SpreadsheetApp.create(responseSpreadsheetName);
      studentResponseSpreadsheetFile = DriveApp.getFileById(studentResponseSpreadsheet.getId());
      studentResponseSpreadsheet.addViewer("krank23@gmail.com"); // TODO: Replace when not in testing

      // studentResponseSpreadsheet.addViewer(student.email);
      // studentResponseSpreadsheetFile.addCommenter("krank23@gmail.com");
      // studentResponseSpreadsheetFile.addCommenter(student.email);
    }

    // -- Set folder
    studentResponseSpreadsheetFile.moveTo(targetFolder);

    return studentResponseSpreadsheet;
  }

  //#endregion
}