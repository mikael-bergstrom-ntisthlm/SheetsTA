/// <reference path="../libs/rubrics.ts" />
/// <reference path="../pages/mastergradingsheet.ts" />
/// <reference path="../pages/detailssheet.ts" />


function Test() {
  // let masterSheet = SpreadsheetApp.getActive().getSheetByName("Bedömning");
  // let gradingSheet = SpreadsheetApp.getActive().getSheetByName("_STUDENTGRADE");
  // if (masterSheet == null || gradingSheet == null) return;

  StudentGradingSheetTA.Setup.CreateOrUpdateStudentGradingSheet(
    SpreadsheetApp.getActive(),
    "Bedömning"
  );
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

  interface RangeValuePair {
    range: GoogleAppsScript.Spreadsheet.Range,
    values: any[][]
  }

  export namespace Setup {
    export function CreateOrUpdateStudentGradingSheet(
      spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
      masterGradingSheetName: string
    ) {

      const setup: DetailsTA.SheetSetup = {
        ColRubric: 1,
        ColCriteria:2,
        ColColnum:3,
        ColCheckmark:4,
        ColGrade:5,
        ColActive: 6,
        RowHeader: 3,
        
        IncludeCheckboxCol: true,
        IncludeGradeCol: false,
        IncludeGradeLine: true,
        IncludeCommentLine: true,
    
        CheckboxColType: "checkable",
        CheckboxColColorized: true,
      }

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
      DetailsTA.SetupHeaderBlock(studentGradingSheet, masterGradingSheet, setup);

      SetupRubricsBlock(studentGradingSheet, rubrics);

      // -- SET WIDTHS
      studentGradingSheet
        .setColumnWidth(_ColRubric, 223)
        .setColumnWidth(_ColCriteria, 275)
        .setColumnWidth(_ColGrade, 70)
        .setColumnWidth(_ColActive, 70)
        .hideColumns(_ColColnum);

    }

    // TODO: Generalize tis, so can be reused at least partly for _TEMPLATE and _VIEW
    function SetupRubricsBlock(
      studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
      rubrics: RubricsTA.Rubric[]
    ) {

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
        dataValues[row][_ColActive - 1] = true;

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

    function FormatRubricBlock(
      rubric: RubricsTA.Rubric,
      studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
      rubricBlockStartRow: number
    ) {

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

  function GetSyncDataPairs(
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    userId: string
  ): {masterData: RangeValuePair, gradingData: RangeValuePair } | null {
    // -- PREP RANGES & VALUES

    // Get target range & values
    const masterDataRange = MasterGradingSheetTA.GetStudentData(userId, masterGradingSheet);

    if (!masterDataRange?.dataRange) {
      Browser.msgBox("User ID not found!");
      return null;
    }

    const masterData: RangeValuePair = {
      range: masterDataRange.dataRange,
      values: masterDataRange.dataRange.getValues()
    }

    const targetStudentValues = masterDataRange.dataRange?.getValues();
    if (!targetStudentValues) return null;

    // Get grading data range & values
    const gradingData = GetGradingData(studentGradingSheet);

    return { masterData, gradingData }
  }

  function GetGradingData(studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): RangeValuePair {

    const gradingDataRange = studentGradingSheet
      .getRange(_RowHeader + 1, 1, studentGradingSheet.getLastRow() - _RowHeader,
        Math.max(_ColActive, _ColCheckmark, _ColColnum, _ColCriteria, _ColGrade, _ColRubric));

    return {
      values: gradingDataRange.getValues(),
      range: gradingDataRange
    };
  }

  export function TransferToMasterSheet(
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    userId: string,
    clearAfterTransfer: boolean
  ) {

    let pairs = GetSyncDataPairs(masterGradingSheet, studentGradingSheet, userId);
    if (!pairs) return;
    const { masterData, gradingData } = pairs;

    // -- PROCESS
    let overrideChecked: boolean = false;
    let cancelled = gradingData.values?.some((row, rowNum) => {

      let targetColumnNum = parseInt(row[_ColColnum - 1]);
      if (isNaN(targetColumnNum)) return;

      if (!(masterData.values[0][targetColumnNum].length == 0) && !overrideChecked)
      {
        let answer = Browser.msgBox(
        "Warning!",
          "Grading data for student already exists.Overwrite ? ",
          Browser.Buttons.YES_NO
        );
        if (answer === "no") return true;
        overrideChecked = true;
      }

      // Transfer data point
      masterData.values[0][targetColumnNum] = row[_ColCheckmark - 1];

      // Reset?
      if (clearAfterTransfer) {
        if (row[_ColCheckmark - 1] === "✔" || row[_ColCheckmark - 1] === "✘") row[_ColCheckmark - 1] = ["✘"]
        else row[_ColCheckmark - 1] = "";

        gradingData.values[rowNum] = row;
      }
    });

    if (cancelled) return;

    // -- POST-PROCESS
    masterData.range?.setValues(masterData.values);

    if (clearAfterTransfer) {
      gradingData.range.setValues(gradingData.values);
      ClearSelectedUserId(studentGradingSheet);
    }
  }

  export function ImportFromMasterSheet(
    masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    studentGradingSheet: GoogleAppsScript.Spreadsheet.Sheet,
    userId: string
  ) {

    let pairs = GetSyncDataPairs(masterGradingSheet, studentGradingSheet, userId);
    if (!pairs) return;
    const { masterData, gradingData } = pairs;

    // -- PROCESS
    gradingData.values?.forEach((row, rowNum) => {

      let sourceColumnNum = parseInt(row[_ColColnum - 1]);
      if (isNaN(sourceColumnNum)) return;

      gradingData.values[rowNum][_ColCheckmark - 1] =
        masterData.values[0][sourceColumnNum]
    });

    // -- POST-PROCESS
    gradingData.range.setValues(gradingData.values);
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