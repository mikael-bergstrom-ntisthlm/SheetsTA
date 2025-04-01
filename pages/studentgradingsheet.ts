/// <reference path="../libs/rubrics.ts" />
/// <reference path="../pages/mastergradingsheet.ts" />

function Test() {
  // let sheet = SpreadsheetApp.getActive().getSheetByName("Bedömning");
  // if (!sheet) return;

  StudentGradingSheetTA.Setup.CreateOrUpdateStudentGradingSheet(SpreadsheetApp.getActive(), "Bedömning");

  // let result = MasterGradingSheetTA.GetStudentData("105003234631509491556", sheet);


  // Logger.log(result);
}

namespace StudentGradingSheetTA {

  const _ColRubric: number = 1;
  const _ColCriteria: number = 2;
  const _ColColnum: number = 3;
  const _ColCheckmark: number = 4;
  const _ColGrade: number = 5;
  const _ColActive: number = 6;

  const _RowHeader: number = 3;
  const _EditBoxColor: number[] = [217, 234, 211];

  export namespace Setup {
    export function CreateOrUpdateStudentGradingSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, masterGradingSheetName: string) {

      // -- PREP
      const masterGradingSheet = spreadsheet.getSheetByName(masterGradingSheetName);
      if (!masterGradingSheet) return;

      const studentGradingSheet = SheetsTA.CreateOrGetSheet("_STUDENTGRADE", spreadsheet, true);
      if (!studentGradingSheet) return;

      const rubrics = RubricsTA.GetRubrics(masterGradingSheet);

      // -- CLEAR & SET SIZE
      SheetsTA.ClearSheet(studentGradingSheet);

      const totalHeight = _RowHeader
        + RubricsTA.CountCriteria(rubrics)
        + rubrics.length * 2 // Space for grade + spacing
        + 3; // Space for comment block
      
      SheetsTA.SetSheetSize(studentGradingSheet, 8, totalHeight);

      // -- SETUP BLOCKS
      SetupHeaderBlock(studentGradingSheet, masterGradingSheet);
      SetupRubricsBlock(studentGradingSheet, rubrics);

      // -- SET WIDTHS
      studentGradingSheet
        .setColumnWidth(_ColRubric, 223)
        .setColumnWidth(_ColCriteria, 275)
        .setColumnWidth(_ColGrade, 70)
        .setColumnWidth(_ColActive, 70)
        .hideColumns(_ColColnum);
      
    }

    function SetupHeaderBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {

      // -- PREP
      const studentNameIds: string[] = MasterGradingSheetTA.GetStudentsData(masterGradingSheet)
        .map(student => student.name + " " + student.surname + " | " + student.id);

        
        // 3 header rows: Student choice, blank, headings
        const headerRange = studentGradingSheet.getRange(1, 1, _RowHeader, 6);
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
        studentGradingSheet.setFrozenRows(_RowHeader);
    }

    function SetupRubricsBlock(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet, rubrics: RubricsTA.Rubric[]) {

      let rubricStartRow = _RowHeader + 1;

      const dataRange = studentGradingSheet.getRange(rubricStartRow, 1, studentGradingSheet.getMaxRows() - _RowHeader, 6);
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

      // -- COMMENT ROW
      dataValues[row + 1][_ColCriteria - 1] = "Comment";

      // Offset is 3 because last criteria's colnr + last grade colnr + 2.
      const commentColNr = 3 + (rubrics.at(-1)?.criteria.at(-1)?.columnNumber ?? 0);
      dataValues[row + 1][_ColColnum - 1] = commentColNr.toString();

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
  }

  export function ClearGrading(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const checkmarkRange = studentGradingSheet.getRange(
      _RowHeader + 1,
      _ColCheckmark,
      studentGradingSheet.getMaxRows() - _RowHeader + 1 // TODO: How to improve speed?
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

    // TODO: Implement copying of comment cell

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