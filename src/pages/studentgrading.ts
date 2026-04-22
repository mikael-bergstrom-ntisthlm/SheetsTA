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
  const _EditBoxColor: number[] = [217, 234, 211];

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
        RowHeader: _RowHeader,

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

      const gradingOverviewSheet = PageGradingOverview.GetDefaultGradingOverviewSheet(spreadsheet);
      const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);

      if (!studentGradingSheet || !gradingOverviewSheet || !rubricsSheet) {
        SpreadsheetApp.getUi().alert("At least one sheet not found (student grading, overview, rubrics)");
        return;
      }

      // -- GET DATA

      const rubrics = PageRubrics.GetRubrics(rubricsSheet);
      const students = PageGradingOverview.GetStudentsData(gradingOverviewSheet);

      // -- CLEAR & SET SIZE
      LibGSheets.ClearSheet(studentGradingSheet);

      const totalHeight = setup.RowHeader
        + LibRubrics.CountCriteria(rubrics)
        + rubrics.length * 2 // Space for grade + spacing
        + 3; // Space for comment block

      const totalWidth = PageStudentDetails.GetHighestColumnNumber(setup) + 1;

      LibGSheets.SetSheetSize(studentGradingSheet, totalWidth, totalHeight);

      // -- SETUP BLOCKS

      PageStudentDetails.SetupHeaderBlock(studentGradingSheet, students, setup);
      SetupRubricsBlock(studentGradingSheet, rubrics, setup);

      // -- SET WIDTHS
      studentGradingSheet
        .setColumnWidth(_ColRubric, 223)
        .setColumnWidth(_ColCriteria, 275)
        .setColumnWidth(_ColGrade, 70)
        .setColumnWidth(_ColActive, 70)
      studentGradingSheet
        .hideColumns(_ColTag);

    }

    /**
     * Add a block of rubrics & criteria
     * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The sheet to add rubrics block to
     * @param {LibRubrics.Rubric[]} rubrics - The rubrics to add rows etc for
     */
    export function SetupRubricsBlock(
      studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
      rubrics: LibRubrics.Rubric[],
      setup: PageStudentDetails.SheetSetup
    ) {
      //TODO: Is the plan to move this too to studentdetails?

      const rubricStartRow = _RowHeader + 1;
      const width = PageStudentDetails.GetHighestColumnNumber(setup);

      const dataRange = studentGradingSheet.getRange(
        rubricStartRow, 1,
        studentGradingSheet.getMaxRows() - _RowHeader,
        width);
      const dataValues = dataRange.getValues();

      // -- RUBRICS ROWS
      let row = 0;

      rubrics.forEach(rubric => {
        let rubricBlockStartRow = rubricStartRow + row;

        dataValues[row][0] = rubric.name;

        // Insert rows from criteria
        rubric.criteria.forEach(criteria => {
          dataValues[row][setup.ColCriteria - 1] = criteria.name;
          dataValues[row][setup.ColTag - 1] = criteria.tag;
          dataValues[row][setup.ColCheckmark - 1] = "✘";
          dataValues[row][setup.ColGrade - 1] = criteria.grade;
          dataValues[row][setup.ColActive - 1] = criteria.active;
          row++;
        });

        // "Grade" on its own row
        dataValues[row][setup.ColCriteria - 1] = "Grade";
        dataValues[row][setup.ColTag - 1] = rubric.gradeTag;
        dataValues[row][setup.ColActive - 1] = true;

        row += 2;

        // When done, format the block
        FormatRubricBlock(studentGradingSheet, rubric.criteria.length, rubricBlockStartRow)
      });

      // -- COMMENT ROW
      dataValues[row + 1][setup.ColCriteria - 1] = "Comment";
      dataValues[row + 1][setup.ColTag - 1] = "comment";

      dataRange.offset(row + 1, setup.ColCriteria - 1, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold")
        .offset(0, 2, 1, 3) // get writing box
        // TODO: Four magic numbers; not ideal
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
