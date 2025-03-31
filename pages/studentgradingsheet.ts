/// <reference path="../libs/rubrics.ts" />
/// <reference path="../pages/mastergradingsheet.ts" />

namespace StudentGradingSheetTA {

  const _ColRubric: number = 1;
  const _ColCriteria: number = 2;
  const _ColColnum: number = 3;
  const _ColCheckmark: number = 4;
  const _ColGrade: number = 5;
  const _ColActive: number = 6;

  const _HeaderRow: number = 3;
  const _EditBoxColor: number[] = [217, 234, 211];

  export namespace Setup {
    export function CreateOrUpdateStudentGradingSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, masterGradingSheetName: string) {

      // -- PREP
      const masterGradingSheet = spreadsheet.getSheetByName(masterGradingSheetName);
      if (!masterGradingSheet) return;

      const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", spreadsheet, true);
      if (!studentGradingSheet) return;

      // -- CLEAR & CLEAN
      studentGradingSheet.clear();
      studentGradingSheet.getFilter()?.remove();

      studentGradingSheet.getRange(1, 1,
        studentGradingSheet.getMaxRows(),
        studentGradingSheet.getMaxColumns())
        .removeCheckboxes()
        .getMergedRanges().forEach(mergedRange => mergedRange.breakApart());

      // TODO: Remove checkmarks, allow page to grow

      // -- SETUP BLOCKS
      SetupHeaderBlock(studentGradingSheet, masterGradingSheet);
      SetupRubricsBlock(studentGradingSheet, masterGradingSheet);

      // -- SET WIDTHS
      studentGradingSheet
        .setColumnWidth(_ColRubric, 223)
        .setColumnWidth(_ColCriteria, 275)
        .setColumnWidth(_ColGrade, 70)
        .setColumnWidth(_ColActive, 70)
        .hideColumns(_ColColnum);

      // -- TRIMMING
      TrimSheetToContents(studentGradingSheet, 1);
    }

    function SetupHeaderBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, rosterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

      const studentNameIds = MasterGradingSheetTA.GetStudentNameIds(rosterGradingSheet);

      studentGradingSheet.setFrozenRows(_HeaderRow);

      // 3 header rows: Student choice, blank, headings
      const headerRange = studentGradingSheet.getRange(1, 1, studentGradingSheet.getMaxRows(), 6);
      const headerValues = headerRange.getValues();

      // Setup student name cells
      headerValues[0][0] = "Student name:";
      studentGradingSheet.getRange(1, 2, 1, 3).merge();

      let rule = SpreadsheetApp.newDataValidation().requireValueInList(studentNameIds).build();
      studentGradingSheet.getRange(1, 2)
        .setDataValidation(rule);

      // Setup data headers
      headerValues[2][_ColRubric - 1] = "Rubric";
      headerValues[2][_ColCriteria - 1] = "Criteria";
      headerValues[2][_ColColnum - 1] = "Column number";
      headerValues[2][_ColCheckmark - 1] = "Check";
      headerValues[2][_ColGrade - 1] = "Grade";
      headerValues[2][_ColActive - 1] = "Active";

      headerRange.setValues(headerValues);
    }

    function SetupRubricsBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
      const rubrics = RubricsTA.GetRubrics(masterGradingSheet);

      let rubricStartRow = _HeaderRow + 1;

      const dataRange = studentGradingSheet.getRange(rubricStartRow, 1, studentGradingSheet.getMaxRows(), 6);
      const dataValues = dataRange.getValues();

      // Insert rows from rubrics
      let row = 0;

      rubrics.forEach(rubric => {
        let rubricBlockStartRow = rubricStartRow + row;

        dataValues[row][0] = rubric.name;

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

        row += 2;

        FormatRubricBlock(rubric, studentGradingSheet, rubricBlockStartRow)
      });

      // Comment row
      dataValues[row + 1][_ColCriteria - 1] = "Comment";
      // TODO: Target column number for comment

      dataRange.offset(row + 1, _ColCriteria - 1, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold")
        .offset(0, 2, 1, 3) // get writing box
        .setBackgroundRGB(_EditBoxColor[0], _EditBoxColor[1], _EditBoxColor[2])
        .merge();


      dataRange.setValues(dataValues);

      // General formatting
      dataRange.setWrap(true);
      dataRange.setVerticalAlignment("top");

      // ADD FILTER
      SetFilter(rubrics, dataRange);
    }

    function FormatRubricBlock(rubric: RubricsTA.Rubric, studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, rubricBlockStartRow: number) {

      // Rubric label block
      studentGradingSheet.getRange(rubricBlockStartRow, _ColRubric, rubric.criteria.length + 1, 1)
        .merge()
        .setBackground("#EFEFEF")
        .setFontWeight("bold");

      // Checkboxes
      studentGradingSheet.getRange(rubricBlockStartRow, _ColCheckmark, rubric.criteria.length, 1)
        .setHorizontalAlignment("center")
        .insertCheckboxes("✔", "✘");

      // Grade sub-block
      studentGradingSheet.getRange(rubricBlockStartRow + rubric.criteria.length, _ColCriteria, 1, 1)
        .setHorizontalAlignment("right")
        .setFontWeight("bold");

      studentGradingSheet.getRange(rubricBlockStartRow + rubric.criteria.length, _ColCheckmark, 1, 1)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackgroundRGB(_EditBoxColor[0], _EditBoxColor[1], _EditBoxColor[2])
    }

    function SetFilter(rubrics: RubricsTA.Rubric[], dataRange: GoogleAppsScript.Spreadsheet.Range) {
      // Count number of criteria
      const totalHeight = rubrics.reduce(
        (accumulator, rubric) => {
          return accumulator + rubric.criteria.length;
        }, 0
      )
        + rubrics.length * 2; // Add 1 for the grade and 1 for spacing, for each rubric

      let filterRange = dataRange.offset(-1, 0, totalHeight);
      let filter = filterRange.createFilter();
      const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["FALSE"]);
      filter.setColumnFilterCriteria(_ColActive, criteria);
    }

    function TrimSheetToContents(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, margin: number) {
      const lastColumn = studentGradingSheet.getLastColumn();
      const lastRow = studentGradingSheet.getLastRow();

      const currentColumns = studentGradingSheet.getMaxColumns();
      const currentRows = studentGradingSheet.getMaxRows();

      if (currentColumns > lastColumn + margin)
        studentGradingSheet.deleteColumns(
          lastColumn + margin,
          currentColumns - lastColumn - margin
        );

      if (currentRows > lastRow + margin)
        studentGradingSheet.deleteRows(
          lastRow + margin,
          currentRows - lastRow - margin
        );
    }
  }

  export function Clear(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const checkmarkRange = studentGradingSheet.getRange(
      _HeaderRow + 1,
      _ColCheckmark,
      studentGradingSheet.getMaxRows() - _HeaderRow + 1 // TODO: How to improve speed?
    )

    const checkmarkValues = checkmarkRange.getValues().map(row => {
      if (row[0] === "✔" || row[0] === "✘") return ["✘"]
      else return [""];
    });

    checkmarkRange.setValues(checkmarkValues);

    ClearSelectedUserId(studentGradingSheet);
  }

  export function TransferToMasterSheet(
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    clearAfterTransfer: boolean) {

    let userId = GetSelectedUserId(studentGradingSheet);
    if (userId === "") return;

    // -- PREP RANGES & VALUES
    const targetStudentData = MasterGradingSheetTA.GetStudentData(userId, masterGradingSheet);

    if (targetStudentData == null) {
      Browser.msgBox("User ID not found!");
      return;
    }

    const targetStudentValues = targetStudentData.dataRange?.getValues();
    if (!targetStudentValues) return;

    const studentGradingFilterRange = studentGradingSheet.getFilter()?.getRange();
    if (!studentGradingFilterRange) return;

    const studentGradingFilterValues = studentGradingFilterRange.getValues();

    // -- PROCESS
    studentGradingFilterValues?.forEach((row, rowNum) => {
      let targetColumnNum = parseInt(row[_ColColnum - 1]);

      if (isNaN(targetColumnNum)) return;

      targetStudentValues[0][targetColumnNum] = row[_ColCheckmark - 1];

      if (clearAfterTransfer) {
        if (row[_ColCheckmark - 1] === "✔" || row[_ColCheckmark - 1] === "✘") row[_ColCheckmark - 1] = ["✘"]
        else row[_ColCheckmark - 1] = "";
      }

      if (clearAfterTransfer) studentGradingFilterValues[rowNum] = row;
    });

    // -- POST-PROCESS
    targetStudentData.dataRange?.setValues(targetStudentValues);
    if (!targetStudentData) return;

    if (clearAfterTransfer) {
      studentGradingFilterRange.setValues(studentGradingFilterValues); // Jesus christ is this really necessary
      ClearSelectedUserId(studentGradingSheet);
    }
  }

  export function ImportFromMasterSheet(userId: string,
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    // -- PREP RANGES & VALUES
    const sourceStudentData = MasterGradingSheetTA.GetStudentData(userId, masterGradingSheet);
    if (sourceStudentData == null) {
      Browser.msgBox("User ID not found!");
      return;
    }

    const sourceStudentValues = sourceStudentData?.dataRange?.getValues()[0];
    if (!sourceStudentValues) return;

    const studentGradingFilterRange = studentGradingSheet.getFilter()?.getRange();
    if (!studentGradingFilterRange) return;

    const studentGradingFilterValues = studentGradingFilterRange.getValues();

    // -- PROCESS
    studentGradingFilterValues?.forEach((row, rowNum) => {
      let sourceColumnNum = parseInt(row[_ColColnum - 1]);

      if (isNaN(sourceColumnNum)) return;

      studentGradingFilterValues[rowNum][_ColCheckmark - 1] =
        sourceStudentValues[sourceColumnNum]
    });

    // -- POST-PROCESS
    studentGradingFilterRange.setValues(studentGradingFilterValues);
  }

  export function GetSelectedUserId(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): string {
    const nameCellValue: string = studentGradingSheet.getRange(1, 2).getValue();

    if (nameCellValue == "") {
      Browser.msgBox("No selection!");
      Logger.log("No selection!");
      return "";
    }

    let pair = nameCellValue.split("|");
    if (pair.length != 2 || pair[1] === "") {
      Browser.msgBox("Invalid selection!");
      Logger.log("Invalid selection!");
      return "";
    }

    return pair[1].trim();
  }

  function ClearSelectedUserId(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    studentGradingSheet.getRange(1, 2).setValue("");
  }
}