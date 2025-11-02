import { PageGradingOverview } from "./gradingoverview";

export namespace PageStudentDetails {

  export interface SheetSetup {
    ColRubric: number;
    ColCriteria: number;
    ColColnum: number;
    ColCheckmark: number;
    ColGrade: number;
    ColActive: number;
    RowHeader: number;

    IncludeCheckboxCol: boolean;
    IncludeGradeCol: boolean;
    IncludeGradeLine: boolean;
    IncludeCommentLine: boolean;

    CheckboxColType: "checkable" | "uncheckable" | "placeholder"
    CheckboxColColorized: boolean;
  }

  export function SetupHeaderBlock(
    targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
    students: PageGradingOverview.StudentData[],
    setup: SheetSetup
  ) {
    // -- PREP
    const studentNameIds: string[] = students
      .map(student => student.name + " " + student.surname + " | " + student.id);


    // 3 header rows: Student choice, blank, headings
    const headerRange = targetSheet.getRange(1, 1, setup.RowHeader, 6); // TODO: Get rid of the 6 (width)
    const headerValues = headerRange.getValues();

    // Setup student name cells (Always B1:D1)
    headerValues[0][0] = "Student name:";
    targetSheet.getRange(1, 2, 1, 3).merge();

    let rule = SpreadsheetApp.newDataValidation().requireValueInList(studentNameIds).build();
    targetSheet.getRange(1, 2)
      .setDataValidation(rule);

    // Setup data headers
    headerValues[setup.RowHeader - 1][setup.ColRubric - 1] = "Rubric";
    headerValues[setup.RowHeader - 1][setup.ColCriteria - 1] = "Criteria";
    headerValues[setup.RowHeader - 1][setup.ColColnum - 1] = "Column number";
    if (setup.IncludeCheckboxCol) {
      headerValues[setup.RowHeader - 1][setup.ColCheckmark - 1] = "✔/✘";
      targetSheet.getRange(setup.RowHeader, setup.ColCheckmark).setHorizontalAlignment("center");
    }
    if (setup.IncludeGradeCol)
      headerValues[setup.RowHeader - 1][setup.ColGrade - 1] = "Grade";
    headerValues[setup.RowHeader - 1][setup.ColActive - 1] = "Active";

    headerRange.setValues(headerValues);
    targetSheet.setFrozenRows(setup.RowHeader);
  }

}