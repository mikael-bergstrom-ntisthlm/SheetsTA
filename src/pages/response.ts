import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js";
import { LibStudents } from "../libs/students.js";
import { PageRubrics } from "./rubrics.js";
import { PageStudentDetails } from "./studentdetails.js";

// TODO: Adding classroom & assignment name to filenames of response docs / folder

export namespace PageResponse {

  const _ResponseTemplateSheetName = "_TEMPLATERESPONSE";
  const _ResponseSheetDetailsName = "DETAILS";

  export const _ResponseDocTag = "responsedoc";

  // -- CONFIG
  export const setup: PageStudentDetails.StudentDetailsSetup = {
    ColRubric: 2,
    ColCriteria: 3,
    ColTag: 4,
    ColGrade: 5, // TODO: Make this *optional*; maybe always show it but hide if unwanted
    ColCheckmark: 6,
    ColActive: 7,
    ColHeaderData: 3,
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

    const rubrics = PageRubrics.GetRubrics(rubricsSheet);

    // -- INIT PAGE

    // Clear & set the initial size to "enough"
    LibGSheets.ClearSheet(responseTemplate);
    LibGSheets.SetSheetSize(responseTemplate,
      PageStudentDetails.GetHighestColumnNumber(setup) * 2,
      rubrics.flatMap(rubric => rubric.criteria).length * 2
    );

    // -- SETUP BLOCKS
    PageStudentDetails.SetupHeaderBlock(
      responseTemplate, [],
      setup
    )

    PageStudentDetails.SetupRubricsBlock(
      responseTemplate,
      rubrics,
      setup
    )

    // -- VISUALS

    const checkmarkRange = responseTemplate.getRange(setup.RowHeaderHeight + 1, setup.ColCheckmark, responseTemplate.getLastRow() - setup.RowHeaderHeight);
    const checkmarkGreenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("✔")
      .setBackground('#00ff00')
      .setRanges([checkmarkRange])
      .build();

    const rules = [checkmarkGreenRule];
    responseTemplate.setConditionalFormatRules(rules);

    // TODO: Red backgrounds for empty E-level criterias?

    // -- HIDE TAG COLUMN
    if (setup.ColTag > 0) {
      responseTemplate.hideColumns(setup.ColTag);
    }

    // -- HIDE GRADE COLUMN
    if (setup.ColGrade) {
      responseTemplate.hideColumns(setup.ColGrade)
    }

    // -- SET WIDTHS

    responseTemplate.setColumnWidth(1, 20); // TODO: Do this for student grading too

    if (setup.ColRubric > 0) {
      responseTemplate
        .setColumnWidth(setup.ColRubric, 223)
    }
    if (setup.ColCriteria > 0) {
      responseTemplate
        .setColumnWidth(setup.ColCriteria, 275);
    }
    if (setup.ColGrade > 0) {
      responseTemplate
        .setColumnWidth(setup.ColGrade, 70)
    }


    LibGSheets.TrimSheet(responseTemplate, 1);

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

  /**
   * 
   * @param rowBlocks An array of Ranges; expected to already be full-width
   * @param tagColNumbers A map of tags and columns matching the rowBlocks
   * @param targetFolder The folder where response documents are moved to
   * @param responseTemplateSheet The sheet to use as template
   * @param rubricsSheet The sheet containing rubric data
   * @returns 
   */
  export function GenerateOrUpdateResponseDocuments(
    rowBlocks: GoogleAppsScript.Spreadsheet.Range[],
    tagColNumbers: Map<string, number>,
    studentColumnSetup: LibStudents.StudentColumnSetup,
    targetFolder: GoogleAppsScript.Drive.Folder,
    responseTemplateSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet,
    assignmentName: string
  ) {

    // -- PREPARE

    // -- Get parent spreadsheet; for toasts
    const spreadsheet = responseTemplateSheet.getParent();

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

        // TODO: It would be nice to be able to "force" a *complete* refresh of the target doc


        let studentResponseSpreadsheet =
          GetOrCreateStudentResponseSpreadsheet(student, assignmentName, responseDocUrl, targetFolder);
        if (studentResponseSpreadsheet === undefined) return;

        // Get the right sheet, if it exists
        let responseSheet = PageResponse.GetOrCreateDetailsSheet(
          studentResponseSpreadsheet,
          responseTemplateSheet
        );

        // -- COMBINE STUDENT DATA WITH RESPONSE DOC
        PageStudentDetails.InsertStudentDataRubrics(student, responseSheet, PageResponse.setup);

        // TODO: Support for more (custom) fields; tied to tags
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

  export function ReadDataBackFromResponseDocs(
    rowBlocks: GoogleAppsScript.Spreadsheet.Range[],
    tagColNumbers: Map<string, number>, // of the rowBlocks
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {

    // -- PREP

    // -- Make sure there's a column for the response doc URL
    const responseColNum = tagColNumbers.get(_ResponseDocTag);
    if (responseColNum === undefined) {
      Browser.msgBox(`No response document column found! \\nNeeds to have the tag ${_ResponseDocTag}`);
      return;
    }

    const commentColNum = tagColNumbers.get("comment") ?? -1;

    const rubrics = PageRubrics.GetRubrics(rubricsSheet);


    // -- PROCESS

    // -- Go through all blocks
    rowBlocks.forEach(rowBlock => {

      // Read current block's values
      const rowBlockValues = rowBlock.getValues();

      // Go through the block's rows
      for (let rowNum = 0; rowNum < rowBlockValues.length; rowNum++) {

        // -- Get the response spreadsheet/sheet
        let studentResponseSpreadsheet: undefined | GoogleAppsScript.Spreadsheet.Spreadsheet = undefined;
        const responseUrl = rowBlockValues[rowNum][responseColNum];

        try {
          studentResponseSpreadsheet = SpreadsheetApp.openByUrl(responseUrl);
        } catch {
          continue;
        }

        const responseSheet = studentResponseSpreadsheet.getSheetByName(_ResponseSheetDetailsName);
        if (responseSheet === null) {
          continue;
        }

        // -- Get the student grading data from the response sheet

        const gradingData = PageStudentDetails.GetStudentGradingData(
          rubrics,
          responseSheet,
          setup
        );

        // -- Insert the student grading data into the row

        //TODO: Use InsertGradingDataIntoValuesRow instead?

        const allCriteria = LibRubrics.GetAllCriteria(gradingData.rubrics);

        allCriteria.forEach(criterion => {
          const colNum = tagColNumbers.get(criterion.tag);
          if (colNum === undefined) return;

          rowBlockValues[rowNum][colNum] = criterion.studentPassed ? "✔" : "✘";
        });

        // -- Insert the comment
        if (commentColNum > 0) {
          rowBlockValues[rowNum][commentColNum] = gradingData.comment;
        }
      }

      // -- Write back the values into the block
      rowBlock.setValues(rowBlockValues);

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
    assignmentName: string,
    responseDocUrl: string,
    targetFolder: GoogleAppsScript.Drive.Folder
  ): GoogleAppsScript.Spreadsheet.Spreadsheet | undefined {

    const responseSpreadsheetName = `Response for ${assignmentName}: ${student.surname} ${student.name}`;

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
      } catch {
        const overwrite = Browser.msgBox(`Student "${student.name} ${student.surname} has something in the response doc column, but it doesn't seem to be the url of a Spreadsheet document\\nDo you want to overwrite this content?"`, Browser.Buttons.YES_NO);
        if (overwrite === "no") return undefined;
      }
    }

    if (studentResponseSpreadsheet === undefined || studentResponseSpreadsheetFile === undefined) {
      studentResponseSpreadsheet = SpreadsheetApp.create(responseSpreadsheetName);
      studentResponseSpreadsheetFile = DriveApp.getFileById(studentResponseSpreadsheet.getId());
      try {
        studentResponseSpreadsheet.addViewer(student.email);
      } catch {
        Browser.msgBox(`${student.email} is not a valid E-mail`);
      }
    }

    // -- Set folder
    studentResponseSpreadsheetFile.moveTo(targetFolder);

    return studentResponseSpreadsheet;
  }

  //#endregion
}