import { LibRubrics } from "../libs/rubrics.js";
import { LibGSheets } from "../libs/sheets.js";
import { LibStudents } from "../libs/students.js";
import { PageRubrics } from "./rubrics.js";

export namespace PageStudentDetails {

  const _EditBoxColor: string = "#D9EAD3";

  export function SetupHeaderBlock(
    targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
    students: LibStudents.StudentData[],
    setup: SheetSetup
  ) {
    // -- PREP
    const studentNameIds: string[] = students
      .map(student => student.name + " " + student.surname + " | " + student.id);

    // Header rows
    const headerRange = targetSheet.getRange(1, 1, setup.RowHeaderHeight, GetHighestColumnNumber(setup));
    const headerValues = headerRange.getValues();

    // Setup student name cells
    if (setup.RowHeaderName > 0) {

      headerValues[setup.RowHeaderName - 1][0] = "Student name:";
      targetSheet.getRange(setup.RowHeaderName, 2, 1, 3).merge();

      if (studentNameIds.length > 0) {
        let rule = SpreadsheetApp.newDataValidation().requireValueInList(studentNameIds).build();
        targetSheet.getRange(setup.RowHeaderName, 2)
          .setDataValidation(rule);
      }
    }

    // Setup comment row in header (if any)
    if (setup.RowHeaderComment > 0) {
      headerValues[setup.RowHeaderComment - 1][0] = "Comment:";
      targetSheet.getRange(setup.RowHeaderComment, 2, 1, 3).merge();
    }

    // Setup data headers
    if (setup.ColRubric > 0)
      headerValues[setup.RowHeaderHeight - 1][setup.ColRubric - 1] = "Rubric";
    if (setup.ColCriteria > 0)
      headerValues[setup.RowHeaderHeight - 1][setup.ColCriteria - 1] = "Criteria";
    if (setup.ColTag > 0)
      headerValues[setup.RowHeaderHeight - 1][setup.ColTag - 1] = "Tag";

    if (setup.ColCheckmark > 0) {
      headerValues[setup.RowHeaderHeight - 1][setup.ColCheckmark - 1] = "✔/✘";
      targetSheet.getRange(setup.RowHeaderHeight, setup.ColCheckmark).setHorizontalAlignment("center");
    }

    if (setup.ColGrade > 0)
      headerValues[setup.RowHeaderHeight - 1][setup.ColGrade - 1] = "Grade";

    if (setup.ColActive > 0)
      headerValues[setup.RowHeaderHeight - 1][setup.ColActive - 1] = "Active";

    headerRange.setValues(headerValues);
    targetSheet.setFrozenRows(setup.RowHeaderHeight);
  }

  /**
   * Add a block of rubrics & criteria
   * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet - The sheet to add rubrics block to
   * @param {LibRubrics.Rubric[]} rubrics - The rubrics to add rows etc for
   */
  export function SetupRubricsBlock(
    targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rubrics: LibRubrics.Rubric[],
    setup: PageStudentDetails.SheetSetup
  ) {

    const rubricStartRow = setup.RowHeaderHeight + 1;
    const width = PageStudentDetails.GetHighestColumnNumber(setup);

    const dataRange = targetSheet.getRange(
      rubricStartRow, 1,
      targetSheet.getMaxRows() - setup.RowHeaderHeight,
      width);
    const dataValues = dataRange.getValues();

    // -- RUBRICS ROWS
    let row = 0;

    rubrics.forEach(rubric => {
      let rubricBlockStartRow = rubricStartRow + row;

      dataValues[row][0] = rubric.name;

      // Insert rows from criteria
      rubric.criteria.forEach(criteria => {
        if (setup.ColCriteria > 0)
          dataValues[row][setup.ColCriteria - 1] = criteria.name;
        if (setup.ColTag > 0)
          dataValues[row][setup.ColTag - 1] = criteria.tag;
        if (setup.ColCheckmark > 0)
          dataValues[row][setup.ColCheckmark - 1] = "✘";
        if (setup.ColGrade > 0)
          dataValues[row][setup.ColGrade - 1] = criteria.grade;
        if (setup.ColActive > 0)
          dataValues[row][setup.ColActive - 1] = criteria.active;
        row++;
      });

      // "Grade" on its own row
      if (setup.GradeForEachRubric) {

        dataValues[row][setup.ColCriteria - 1] = "Grade";
        dataValues[row][setup.ColTag - 1] = rubric.gradeTag;
        dataValues[row][setup.ColActive - 1] = true;
        row += 2;

      } else {
        row++;
      }


      // When done, format the block
      FormatRubricBlock(targetSheet, rubric.criteria, rubricBlockStartRow, setup)
    });

    // -- COMMENT ROW
    if (setup.CommentFooter) {

      dataValues[row + 1][setup.ColCriteria - 1] = "Comment";
      dataValues[row + 1][setup.ColTag - 1] = "comment";

      dataRange.offset(row + 1, setup.ColCriteria - 1, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold")
        .offset(0, 2, 1, 3) // get writing box
        // TODO: Four magic numbers; not ideal
        .setBackground(_EditBoxColor)
        .merge();
    }

    // -- FINALIZING
    dataRange.setValues(dataValues);

    // General formatting
    dataRange.setWrap(true);
    dataRange.setVerticalAlignment("top");

    // ADD FILTER
    SetFilter(rubrics, dataRange, setup);
  }

  /**
 * Add formatting to a rubric's block
 * @param studentGradingSheet - The sheet where the formattin's taking place
 * @param numCriteria - Number of criteria rows
 * @param rubricBlockStartRow - The row where the rubric's block starts
 */
  function FormatRubricBlock(
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    criteria: LibRubrics.Criteria[],
    rubricBlockStartRow: number,
    setup: SheetSetup
  ) {

    // Rubric label block
    const height = criteria.length + (setup.GradeForEachRubric ? 1 : 0);

    studentGradingSheet.getRange(rubricBlockStartRow, setup.ColRubric, height, 1)
      .merge()
      .setBackground("#EFEFEF")
      .setFontWeight("bold");

    // Checkboxes
    studentGradingSheet.getRange(rubricBlockStartRow, setup.ColCheckmark, criteria.length, 1)
      .setHorizontalAlignment("center")
      .insertCheckboxes("✔", "✘");

    // Grade sub-block
    if (setup.GradeForEachRubric) {

      studentGradingSheet.getRange(rubricBlockStartRow + criteria.length, setup.ColCriteria, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold");

      studentGradingSheet.getRange(rubricBlockStartRow + criteria.length, setup.ColCheckmark, 1, 1)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackground(_EditBoxColor);
    }
  }

  /**
   * Add a filter to a rubrics block
   * @param rubrics - The rubrics data
   * @param dataRange - The range where the rubric blocks were added
   */
  function SetFilter(
    rubrics: LibRubrics.Rubric[],
    dataRange: GoogleAppsScript.Spreadsheet.Range,
    setup: SheetSetup
  ) {
    // Count number of criteria
    const totalHeight =
      LibRubrics.CountCriteria(rubrics)
      + rubrics.length * (setup.GradeForEachRubric ? 2 : 1);
    // Add (maybe) 1 for the grade and 1 for spacing, for each rubric

    // Create filter range
    let filterRange = dataRange.offset(-1, 0, totalHeight);
    let filter = filterRange.createFilter();

    // Hide inactive criteria, maybe
    if (setup.ColActive > 0) {
      const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["FALSE"]);
      filter.setColumnFilterCriteria(setup.ColActive, criteria);
    }
  }


  export function GetHighestColumnNumber(setup: SheetSetup) {
    return Math.max(
      setup.ColActive,
      setup.ColCheckmark,
      setup.ColCriteria,
      setup.ColGrade,
      setup.ColRubric,
      setup.ColTag
    );
  }

  /* ---------------------------------------------------------------------------
    TRANSFERRING DATA
  ----------------------------------------------------------------------------*/
  //#region Transferring

  /**
   * Insert Student data from some other source, using criteria tags to match
   * with student detail sheet rows
   * @param student The student data to insert
   * @param studentGradingSheet The sheet to insert it into
   */
  export function InsertStudentDataRubrics(
    student: LibStudents.StudentData,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    setup: SheetSetup
  ): void {

    if (!student.gradingData) {
      Browser.msgBox("Student has no data!");
      return;
    }

    const localData = GetRubricsData(studentGradingSheet, setup);

    // Setup quick index of tags and row numbers for easy lookup
    const tagRowNumbers = new Map<string, number>();

    localData.values.forEach((row, rowNum) => {
      tagRowNumbers.set("" + row[setup.ColTag - 1], rowNum);
    });

    // TODO: Do some checking here

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
        localData.values[rowNum][setup.ColCheckmark - 1] =
          criterion.studentPassed ? "✔" : "✘";
      });

      // Set the grade
      if (setup.GradeForEachRubric) {
        const rowNum = tagRowNumbers.get(rubric.gradeTag);
        if (rowNum === undefined) return;
        localData.values[rowNum][setup.ColCheckmark - 1] = rubric.studentGrade;
      }
    });

    // Set the comment
    if (setup.CommentFooter) {
      const rowNum = tagRowNumbers.get("comment");
      if (rowNum) {
        localData.values[rowNum][setup.ColCheckmark - 1] = student.gradingData.comment;
      }
    }


    // Insert the data
    localData.range.setValues(
      localData.values
    )
  }

  export function GetStudentGradingData(
    rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    setup: SheetSetup
  ): LibStudents.GradingData {

    // Get rubrics from rubrics page
    const data: LibStudents.GradingData = {
      rubrics: PageRubrics.GetRubrics(rubricsSheet),
      comment: ""
    }

    if (data.rubrics.length == 0) { Browser.msgBox("No rubrics found") }

    // Get the local values
    const localData = GetRubricsData(studentGradingSheet, setup);

    // Setup quick index of tags and row numbers for easy lookup
    const tagRowNumbers = new Map<string, number>();

    localData.values.forEach((row, rowNum) => {
      tagRowNumbers.set("" + row[setup.ColTag - 1], rowNum);
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
          localData.values[rowNum][setup.ColCheckmark - 1] == "✔" ? true : false;
      });

      // Set the grade
      const rowNum = tagRowNumbers.get(rubric.gradeTag);
      if (rowNum === undefined) return;
      rubric.studentGrade = localData.values[rowNum][setup.ColCheckmark - 1];
    });

    // -- Get the comment
    const rowNum = tagRowNumbers.get("comment");
    if (rowNum) {
      data.comment = localData.values[rowNum][setup.ColCheckmark - 1];
    }

    // Return the data
    return data;
  }


  /**
   * Get the entire rubrics block (range+values) of a student grading sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} studentGradingSheet - The student grading sheet
   * @returns {RangeValuePair} A value-range pair
   */
  function GetRubricsData(
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    setup: SheetSetup
  ): LibGSheets.RangeValuePair {

    const gradingDataRange = studentGradingSheet
      .getRange(setup.RowHeaderHeight + 1, 1, // Start at the row below the header
        studentGradingSheet.getLastRow() - setup.RowHeaderHeight, // Get all the rows, minus the header
        GetHighestColumnNumber(setup)
      ); // Find the rightmost column

    return {
      values: gradingDataRange.getValues(),
      range: gradingDataRange
    };
  }

  //#endregion

  /* ---------------------------------------------------------------------------
    INTERFACES
  ----------------------------------------------------------------------------*/
  //#region Interfaces
  export interface SheetSetup {
    ColRubric: number;
    ColCriteria: number;
    ColTag: number;
    ColCheckmark: number;
    ColGrade: number;
    ColActive: number;
    RowHeaderHeight: number;
    RowHeaderName: number;
    RowHeaderComment: number;

    CommentFooter: boolean;
    GradeForEachRubric: boolean;
  }
  //#endregion

}