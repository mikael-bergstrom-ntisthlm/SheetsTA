/// <reference path="./sheets.ts" />

function Test() {
  RubricsTA.CreateOrUpdateStudentGradingSheet();
}


namespace RubricsTA {
  
  export function CreateOrUpdateStudentGradingSheet() {
    
    // PREP
    const rosterGradingSheet = SpreadsheetApp.getActive().getSheetByName("Bedömning");
    if (!rosterGradingSheet) return;
    
    const rubrics = GetRubrics(rosterGradingSheet);
    const studentNameIds = GetStudentNameIds(rosterGradingSheet);

    const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", SpreadsheetApp.getActive());
    studentGradingSheet.clear();

    // HEADER BLOCK
    studentGradingSheet.setFrozenRows(3);

    // 3 header rows: Student choice, blank, headings
    // const range = studentGradingSheet.getRange(1, 1, studentGradingSheet.getMaxRows(), 6);
    // const values = range.getValues();

    // values[0][0] = "Student name:";

    // SetColumnHeaders(values);
    // SetColumnWidths(studentGradingSheet);

    // Setup name cell
    // studentGradingSheet.getRange(1, 2, 1, 3).merge()

    // let students = studentNameIds;
    // let rule = SpreadsheetApp.newDataValidation().requireValueInList(students).build();
    // studentGradingSheet.getRange(1, 2)
    //   .setDataValidation(rule);

    const range = studentGradingSheet.getRange(studentGradingSheet.getFrozenRows(), 1, studentGradingSheet.getMaxRows(), 6);
    const values = range.getValues();

    // Insert rows from rubrics
    let row = 0;

    rubrics.forEach(rubric => {
      let rubricBlockHeight = 0;
      let rubricBlockStart = row + 1;

      Logger.log(rubric.name);
      values[row][0] = rubric.name;

      rubric.criteria.forEach(criteria => {
        values[row][1] = criteria.name;
        values[row][2] = criteria.columnNumber;
        values[row][3] = criteria.active;
        values[row][4] = criteria.grade;
        values[row][5] = "✘";
        rubricBlockHeight++;
        row++;
      });

      values[row][4] = "Grade";
      rubricBlockHeight++;
      row++;

      FormatRubricTitleBlock(rubricBlockStart, rubricBlockHeight, studentGradingSheet);

      let r = studentGradingSheet.getRange(rubricBlockStart, 6, rubricBlockHeight - 1, 1);
      r.setHorizontalAlignment("center");
      r.insertCheckboxes("✔", "✘");

      row++;
    })

    
    range.setValues(values);


    // WTF
    SheetsTA.GetColumnNum("UserID", rosterGradingSheet, 5)
    

    
    
    // studentGradingSheet.getRange(1, 2)


    // TODO: Filter
    // TODO: Hide colnr column
    // TODO: Separate header-prep, even at cost of additional call


  }

  function SetColumnHeaders(values: any[][]) {
    values[2][0] = "Rubric";
    values[2][1] = "Criteria";
    values[2][2] = "Column number";
    values[2][3] = "Active";
    values[2][4] = "Grade";
    values[2][5] = "Check";
  }

  function SetColumnWidths(sheet: GoogleAppsScript.Spreadsheet.Sheet)
  {
    sheet.setColumnWidth(1, 223);
    sheet.setColumnWidth(2, 275);
    sheet.setColumnWidth(5, 70);
    sheet.setColumnWidth(6, 70);
    sheet.hideColumns(3);
  }

  function FormatRubricTitleBlock(start: number, height: number, sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    let range = sheet.getRange(start, 1, height, 1);

    range.merge();
    range.setWrap(true);
    range.setVerticalAlignment("top");
    range.setBackground("#EFEFEF");
    range.setFontWeight("bold");
  }

  export function GetRubrics(sheet: GoogleAppsScript.Spreadsheet.Sheet): Rubric[] {

    if (!sheet) return [];

    let headerBlock = sheet.getRange(
      1, sheet.getFrozenColumns() + 1,
      sheet.getFrozenRows(),
      sheet.getLastColumn()
    );

    let headerValues = headerBlock?.getValues();

    if (!headerValues || headerValues?.length == 0) return [];

    let rubricTitleRow = headerValues[0];
    let activeRow = headerValues[1]
    let gradeRow = headerValues[2];
    let shortformRow = headerValues[3]
    let criteriaRow = headerValues[4];

    let rubrics: Rubric[] = [];
    let currentRubric: Rubric | undefined = undefined;

    for (let i = 0; i < rubricTitleRow.length; i++) {
      // Detect rubric start
      if (rubricTitleRow[i] !== "") {
        currentRubric = {
          criteria: [],
          columnNumber: sheet.getFrozenColumns() + i,
          name: rubricTitleRow[i]
        }
        rubrics.push(currentRubric);
      }

      // Detect criteria
      if (gradeRow[i] !== "" && currentRubric) {
        currentRubric.criteria.push(
          {
            name: criteriaRow[i],
            shortform: shortformRow[i],
            active: activeRow[i],
            grade: gradeRow[i],
            columnNumber: sheet.getFrozenColumns() + i
          }
        )
      }
    }

    return rubrics;
  }

  function GetStudentNameIds(rosterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

    const headingRowNumber = rosterGradingSheet.getFrozenRows();
    const dataHeight = rosterGradingSheet.getMaxRows() - rosterGradingSheet.getFrozenRows();

    let colnumId = SheetsTA.GetColumnNum("UserID", rosterGradingSheet, headingRowNumber) // 6
    let colnumSurname = SheetsTA.GetColumnNum("Surname", rosterGradingSheet, headingRowNumber)
    let colnumName = SheetsTA.GetColumnNum("Name", rosterGradingSheet, headingRowNumber)

    let studentIds = rosterGradingSheet.getRange(headingRowNumber + 1, colnumId, dataHeight).getValues();
    let studentSurnames = rosterGradingSheet.getRange(headingRowNumber + 1, colnumSurname, dataHeight).getValues();
    let studentNames = rosterGradingSheet.getRange(headingRowNumber + 1, colnumName, dataHeight).getValues();

    const studentList: string[] = [];

    for (let i = 0; i < studentIds.length; i++) {
      if (studentIds[i][0] === "") continue;

      studentList.push(
        studentNames[i][0] + " " + studentSurnames[i][0] + " | " + studentIds[i][0]
      );
    }

    return studentList;
  }

  interface Rubric {
    name: string;
    studentGrade?: string;
    criteria: Criteria[];
    columnNumber: number;
  }

  interface Criteria {
    name: string;
    shortform: string;
    active: boolean;
    grade: string;
    studentPassed?: boolean;
    columnNumber: number;
  }
} 