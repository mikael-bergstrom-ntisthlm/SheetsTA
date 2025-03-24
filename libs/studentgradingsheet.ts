/// <reference path="rubrics.ts" />

function Test() {
  const masterGradingSheet = SpreadsheetApp.getActive().getSheetByName("Bedömning");
  if (!masterGradingSheet) return;

  const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", SpreadsheetApp.getActive(), false);
  if (!studentGradingSheet) return;

  // StudentGradingSheetTA.CreateOrUpdateStudentGradingSheet(masterGradingSheet);

  // let userId = "115898972864944841723";
  let userId = StudentGradingSheetTA.GetSelectedUserId(studentGradingSheet);

  Logger.log(userId);

  StudentGradingSheetTA.TransferToMasterSheet(userId, masterGradingSheet, studentGradingSheet);

}

namespace StudentGradingSheetTA {

  export function CreateOrUpdateStudentGradingSheet(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    // PREP
    const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", SpreadsheetApp.getActive(), true);
    if (!studentGradingSheet) return;

    // CLEAR & CLEAN
    studentGradingSheet.clear();
    studentGradingSheet.getFilter()?.remove();

    // SETUP BLOCKS
    SetupHeaderBlock(studentGradingSheet, masterGradingSheet);
    SetupRubricsBlock(studentGradingSheet, masterGradingSheet);

    SetColumnWidths(studentGradingSheet);
  }

  function SetupHeaderBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, rosterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const studentNameIds = GetStudentNameIds(rosterGradingSheet);

    studentGradingSheet.setFrozenRows(3);

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
    headerValues[2] = ["Rubric", "Criteria", "Column number", "Check", "Grade", "Active"];

    headerRange.setValues(headerValues);
  }

  function SetupRubricsBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, rosterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const rubrics = RubricsTA.GetRubrics(rosterGradingSheet);

    let rubricStartRow = studentGradingSheet.getFrozenRows() + 1;

    const dataRange = studentGradingSheet.getRange(rubricStartRow, 1, studentGradingSheet.getMaxRows(), 6);
    const dataValues = dataRange.getValues();

    // Insert rows from rubrics
    let row = 0;

    rubrics.forEach(rubric => {
      let rubricBlockStartRow = rubricStartRow + row;

      Logger.log(rubric.name);
      dataValues[row][0] = rubric.name;

      let lastColNr = 0;
      rubric.criteria.forEach(criteria => {
        dataValues[row][1] = criteria.name;
        dataValues[row][2] = criteria.columnNumber;
        dataValues[row][3] = "✘";
        dataValues[row][4] = criteria.grade;
        dataValues[row][5] = criteria.active;
        lastColNr = criteria.columnNumber;
        row++;
      });

      // "Grade" on its own row
      dataValues[row][1] = "Grade";
      dataValues[row][2] = rubric.criteria.slice(-1)[0].columnNumber + 1;

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
    let range = studentGradingSheet.getRange(rubricBlockStartRow, 1, rubric.criteria.length + 1, 1);
    range.merge();
    range.setBackground("#EFEFEF");
    range.setFontWeight("bold");

    // Checkboxes
    let r = studentGradingSheet.getRange(rubricBlockStartRow, 4, rubric.criteria.length, 1);
    r.setHorizontalAlignment("center");
    r.insertCheckboxes("✔", "✘");

    // Grade sub-block
    const gradeTextCell = studentGradingSheet.getRange(rubricBlockStartRow + rubric.criteria.length, 2);

    gradeTextCell.setHorizontalAlignment("right")
      .setFontWeight("bold")
      .offset(0, 2)
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
    filter.setColumnFilterCriteria(6, criteria);
  }

  function SetColumnWidths(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    sheet.setColumnWidth(1, 223);
    sheet.setColumnWidth(2, 275);
    sheet.setColumnWidth(5, 70);
    sheet.setColumnWidth(6, 70);
    sheet.hideColumns(3);
  }

  export function TransferToMasterSheet(userID: string, masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const targetStudentRange = GetStudentRange(userID, masterGradingSheet);
    if (targetStudentRange == null)
    {
      Browser.msgBox("User ID not found!");
      return;
    }

    const targetStudentValues = targetStudentRange.getValues();

    const studentGradingFilterValues = studentGradingSheet.getFilter()?.getRange().getValues();

    studentGradingFilterValues?.forEach(row => {
      let targetColumnNum = parseInt(row[2]);
      if (!isNaN(targetColumnNum)) {
        Logger.log(targetColumnNum);
        targetStudentValues[0][targetColumnNum] = row[3];
      }
    });

    targetStudentRange.setValues(targetStudentValues);

  }

  export function ImportFromMasterSheet(userID: string, masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

  }

  export function GetSelectedUserId(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): string {
    const nameCellValue:string = studentGradingSheet.getRange(1, 2).getValue();

    if (nameCellValue == "")
    {
      Browser.msgBox("No selection!");
      Logger.log("No selection!");
      return "";
    }

    let pair = nameCellValue.split("|");
    if (pair.length != 2 || pair[1] === "")
    {
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

  function GetStudentNameIds(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const headingRowNumber = masterGradingSheet.getFrozenRows();
    const dataHeight = masterGradingSheet.getMaxRows() - masterGradingSheet.getFrozenRows();

    let colnumId = SheetsTA.GetColumnNum("UserID", masterGradingSheet, headingRowNumber) // 6
    let colnumSurname = SheetsTA.GetColumnNum("Surname", masterGradingSheet, headingRowNumber)
    let colnumName = SheetsTA.GetColumnNum("Name", masterGradingSheet, headingRowNumber)

    let studentIds = masterGradingSheet.getRange(headingRowNumber + 1, colnumId, dataHeight).getValues();
    let studentSurnames = masterGradingSheet.getRange(headingRowNumber + 1, colnumSurname, dataHeight).getValues();
    let studentNames = masterGradingSheet.getRange(headingRowNumber + 1, colnumName, dataHeight).getValues();

    const studentList: string[] = [];

    for (let i = 0; i < studentIds.length; i++) {
      if (studentIds[i][0] === "") continue;

      studentList.push(
        studentNames[i][0] + " " + studentSurnames[i][0] + " | " + studentIds[i][0]
      );
    }

    return studentList;
  }
}