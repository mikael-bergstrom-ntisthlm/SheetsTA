/// <reference path="../libs/rubrics.ts" />

function Test() {

  // const spreadsheet = SpreadsheetApp.getActive();
  // const spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1E7SV82BqJbA4Qt0vYGATCxF7pXAT7YniXFmovSKRH30/");

  // Teoriprov
  const spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1J-_WeDj_QVBR_lF-kGIwEYTNXSleRV3RsOlsHjs5nSs/");

  // const masterGradingSheet = spreadsheet.getSheetByName("Bedömning");
  // if (!masterGradingSheet) return;

  // StudentGradingSheetTA.GetStudentNameIds(masterGradingSheet);

  StudentGradingSheetTA.Setup.CreateOrUpdateStudentGradingSheet(spreadsheet, "Bedömning");

  // const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", SpreadsheetApp.getActive(), false);
  // if (!studentGradingSheet) return;

  // let userId = "115898972864944841723";
  // let userId = StudentGradingSheetTA.GetSelectedUserId(studentGradingSheet);


  // StudentGradingSheetTA.TransferToMasterSheet(userId, masterGradingSheet, studentGradingSheet, false);


}

namespace StudentGradingSheetTA {

  const _ColRubric: number = 1;
  const _ColCriteria: number = 2;
  const _ColColnum: number = 3;
  const _ColCheckmark: number = 4;
  const _ColGrade: number = 5;
  const _ColActive: number = 6;

  const _HeaderRow: number = 3;

  export namespace Setup {
    export function CreateOrUpdateStudentGradingSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, masterGradingSheetName: string) {

      // PREP
      const masterGradingSheet = spreadsheet.getSheetByName(masterGradingSheetName);
      if (!masterGradingSheet) return;

      const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", spreadsheet, true);
      if (!studentGradingSheet) return;

      // CLEAR & CLEAN
      studentGradingSheet.clear();
      studentGradingSheet.getFilter()?.remove();

      // SETUP BLOCKS
      SetupHeaderBlock(studentGradingSheet, masterGradingSheet);
      SetupRubricsBlock(studentGradingSheet, masterGradingSheet);

      // SET WIDTHS
      studentGradingSheet
        .setColumnWidth(_ColRubric, 223)
        .setColumnWidth(_ColCriteria, 275)
        .setColumnWidth(_ColGrade, 70)
        .setColumnWidth(_ColActive, 70)
        .hideColumns(_ColColnum);

      // TRIMMING
      TrimSheetToContents(studentGradingSheet, 1);

    }

    function SetupHeaderBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, rosterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

      const studentNameIds = GetStudentNameIds(rosterGradingSheet);

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
        .setBackgroundRGB(217, 234, 211);
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

      studentGradingSheet.deleteColumns(
        lastColumn + margin,
        studentGradingSheet.getMaxColumns() - lastColumn - margin
      );

      studentGradingSheet.deleteRows(
        lastRow + margin,
        studentGradingSheet.getMaxRows() - lastRow - margin
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
  }

  export function TransferToMasterSheet(userID: string,
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    clearAfterTransfer: boolean) {

    const targetStudentRange = GetStudentRange(userID, masterGradingSheet);
    if (targetStudentRange == null) {
      Browser.msgBox("User ID not found!");
      return;
    }

    const targetStudentValues = targetStudentRange.getValues();

    const studentGradingFilterRange = studentGradingSheet.getFilter()?.getRange();
    if (!studentGradingFilterRange) return;

    const studentGradingFilterValues = studentGradingFilterRange.getValues();

    studentGradingFilterValues?.forEach((row, rowNum) => {
      let targetColumnNum = parseInt(row[2]);

      if (!isNaN(targetColumnNum)) {
        targetStudentValues[0][targetColumnNum] = row[3];

        if (clearAfterTransfer) {
          if (row[0] === "✔" || row[0] === "✘") row[3] = ["✘"]
          else row[3] = "";
        }
      }
      if (clearAfterTransfer) studentGradingFilterValues[rowNum] = row;
    });

    targetStudentRange.setValues(targetStudentValues);

    if (clearAfterTransfer) {
      studentGradingFilterRange.setValues(studentGradingFilterValues); // Jesus christ is this really necessary
    }
  }

  export function ImportFromMasterSheet(userID: string,
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    
    const targetStudentRange = GetStudentRange(userID, masterGradingSheet);
    if (targetStudentRange == null) {
      Browser.msgBox("User ID not found!");
      return;
    }

    const targetStudentValues = targetStudentRange.getValues();

    const studentGradingFilterRange = studentGradingSheet.getFilter()?.getRange();
    if (!studentGradingFilterRange) return;

    const studentGradingFilterValues = studentGradingFilterRange.getValues();


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

  // TODO: Move these to module dealing w/ the master grading sheet
  export function GetStudentRange(userID: string, masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): GoogleAppsScript.Spreadsheet.Range | null {
    const headingRowNumber = masterGradingSheet.getFrozenRows();

    // Get the ID column
    let colnumId = SheetsTA.GetColumnNum("UserID", masterGradingSheet, headingRowNumber) // 6
    const idColumnRange = masterGradingSheet.getRange(headingRowNumber + 1, colnumId, masterGradingSheet.getMaxRows());

    // Find the right row
    let IDs = idColumnRange.getValues().map(cell => cell[0]).filter(cell => cell.length > 0);
    let rowNum = IDs.findIndex(id => id === userID);

    if (rowNum < 0) return null;

    const studentRange = masterGradingSheet.getRange(headingRowNumber + rowNum + 1, 1, 1, masterGradingSheet.getMaxColumns())

    return studentRange;
  }

  export function GetStudentNameIds(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const headingRowNumber = masterGradingSheet.getFrozenRows();
    const dataHeight = masterGradingSheet.getMaxRows() - headingRowNumber;

    let studentIds = GetContentsOfColumn("userid", headingRowNumber, dataHeight, masterGradingSheet);
    let studentSurnames = GetContentsOfColumn("surname", headingRowNumber, dataHeight, masterGradingSheet);
    let studentNames = GetContentsOfColumn("name", headingRowNumber, dataHeight, masterGradingSheet);

    const studentList: string[] = [];

    for (let i = 0; i < studentIds.length; i++) {
      if (studentIds[i] === "") continue;

      studentList.push(
        studentNames[i] + " " + studentSurnames[i] + " | " + studentIds[i]
      );
    }

    return studentList;
  }

  function GetContentsOfColumn(columnName: string, headingRowNumber: number, dataHeight: number, sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {

    let colId = SheetsTA.GetColumnNum(columnName, sheet, headingRowNumber);
    if (colId < 0) Browser.msgBox(`Column "${columnName}" not found`);

    let data = sheet.getRange(headingRowNumber + 1, colId, dataHeight)
      .getValues()
      .map(item => item[0]);

    return data;
  }
}