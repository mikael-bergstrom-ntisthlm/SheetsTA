import { LibGClassroom } from "../libs/classroom.js";
import { LibConfig } from "../libs/config.js";
import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js";
import { LibStudents } from "../libs/students.js";
import { PageResponse } from "./response.js";
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

  const _ResponseDocTag = "responsedoc";

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

      fullNameRange.setFormula(`=${String.fromCharCode(64 + _ColSurname)}${_RowDataStart} & " " & ${String.fromCharCode(64 + _ColName)}${_RowDataStart}`);

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
      rubricHeaderRangeValues[_RowTag - 1][responseDocColNum] = _ResponseDocTag;

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

    return GetStudentsDataFromValues(studentsRange.getValues());
  }

  export function GetStudentsDataFromValues(
    sourceValues: any[][]
  ): LibStudents.StudentData[] {

    const studentsData: LibStudents.StudentData[] = [];

    sourceValues.forEach(row => {
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
    data: LibStudents.GradingData,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {

    // -- PREPARE

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

    // Find the first column that contains a criteria
    const allCriteria = data.rubrics.flatMap((rubric) => rubric.criteria);
    const colDataStart = Math.min(...allCriteria.map(criteria => tagColNumbers.get(criteria.tag) ?? 0)) + 1;

    // -- FIND THE RIGHT STUDENT
    const studentsData = GetAllStudentsData(gradingOverviewSheet);

    let studentRowNum = studentsData.findIndex(student => student.id === userID);
    if (studentRowNum < 0) { Browser.msgBox("Student ID not found"); return null; };

    // -- Get the student's data
    const studentData = GetGradingDataRow(
      studentRowNum, colDataStart, gradingOverviewSheet
    );
    

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
      
      // colNumber -= (colDataStart - 1);
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

    InsertRowDataIntoStudent(student, tagColNumbers, studentDataValues);

    return student;
  }

  /**
   * Takes a set of row-data and inserts it into a Student object, using a tag
   * map to determine which of the row's columns maps to which criterion
   * @param student {LibStudents.StudentData}
   * @param tagColNumbers {Map<string, number>}
   * @param studentDataValues {any[][]}
   */
  function InsertRowDataIntoStudent(
    student: LibStudents.StudentData,
    tagColNumbers: Map<string, number>,
    studentDataValues: string[]
  ): void {

    if (student.gradingData === undefined) {
      return;
    }

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
          studentDataValues[colNumber] == "✔";
      });

      // Find column of rubric's overall grade
      const gradeCol = tagColNumbers.get(rubric.gradeTag);
      if (gradeCol === undefined) {
        Browser.msgBox(`No grade column found for rubric "${rubric.name}"`);
      } else {
        // Save rubric grade
        rubric.studentGrade = studentDataValues[gradeCol];
      }
    });

    // Get comment
    const colNumber = tagColNumbers.get("comment");
    if (colNumber === undefined) {
      Browser.msgBox("No column found for comment");
    }
    else {
      student.gradingData.comment = studentDataValues[colNumber];
    }
  }

  // TODO: Determine if this is relevant
  export function GetSelectedStudent(
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {
    const studentDataValues = gradingOverviewSheet.getRange(
      gradingOverviewSheet.getCurrentCell()?.getRow() ?? 0,
      1,
      1,
      gradingOverviewSheet.getMaxColumns()
    ).getValues()[0].map(v => String(v));

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

    // -- Make base student object
    const student: LibStudents.StudentData = {
      id: studentDataValues[_ColUserId - 1],
      name: studentDataValues[_ColName - 1],
      surname: studentDataValues[_ColSurname - 1],
      email: studentDataValues[_ColEmail - 1],
      gradingData: { // Wi
        rubrics: [],
        comment: ""
      }
    }

    Browser.msgBox(student.name);

    InsertRowDataIntoStudent(student, tagColNumbers, studentDataValues)

    // Get current selection
    // Get current row
    // Create student object, return it
  }

  //#endregion

  /* -----------------------------------------------------------------------------
    RESPONSE DOCUMENT GENERATION
  ------------------------------------------------------------------------------*/
  //#region response doc gen

  // TODO: CURRENT PROJECT
  export function GenerateResponseDocuments(
    rowBlocks: GoogleAppsScript.Spreadsheet.Range[],
    targetFolder: GoogleAppsScript.Drive.Folder,
    gradingOverviewSheet: GoogleAppsScript.Spreadsheet.Sheet,
    responseTemplateSheet: GoogleAppsScript.Spreadsheet.Sheet
  ) {

    // -- Make a map of which column belongs to which tag
    const tagColNumbers = MakeTagColNumberMap(gradingOverviewSheet);

    const responseColNum = tagColNumbers.get(_ResponseDocTag);
    if (responseColNum === undefined) {
      Browser.msgBox(`No response document column found! \\nNeeds to have the tag ${_ResponseDocTag}`);
      return;
    }

    rowBlocks.forEach(rowBlock => {
      const rowBlockValues = rowBlock.getValues();
      const students = GetStudentsDataFromValues(rowBlockValues);

      for (let i = 0; i < students.length; i++) {

        const student = students[i];

        let responseDocUrl: string = rowBlockValues[i][responseColNum];

        let studentResponseSpreadsheet =
          GetOrCreateStudentResponseSpreadsheet(student, responseDocUrl, targetFolder);

        if (studentResponseSpreadsheet === undefined) return;

        // Get the right sheet, if it exists
        let sheet = PageResponse.GetOrCreateDetailsSheet(
          studentResponseSpreadsheet,
          responseTemplateSheet
        );



        // If it does not exist, copy template into it
        // Then insert student grading data


        const newUrl = studentResponseSpreadsheet.getUrl();
        if (newUrl !== responseDocUrl) {
          let responseBlock = rowBlock.offset(
            i,
            responseColNum,
            1, 1
          );

          responseBlock.setValue(newUrl)
        }
      }

    });
  }

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
    }

    // -- Set folder
    studentResponseSpreadsheetFile.moveTo(targetFolder);

    return studentResponseSpreadsheet;
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
  function MakeTagColNumberMap(
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

  //#endregion
}