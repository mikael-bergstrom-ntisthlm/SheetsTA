/// <reference path="./sheets.ts" />


namespace RubricsTA {
  export function GetRubrics(masterGradingSheet: GoogleAppsScript.Spreadsheet.Sheet): Rubric[] {

    if (!masterGradingSheet) return [];

    let headerBlock = masterGradingSheet.getRange(
      1, masterGradingSheet.getFrozenColumns() + 1,
      masterGradingSheet.getFrozenRows(),
      masterGradingSheet.getLastColumn()
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
          columnNumber: masterGradingSheet.getFrozenColumns() + i,
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
            columnNumber: masterGradingSheet.getFrozenColumns() + i
          }
        )
      }
    }

    return rubrics;
  }

  export function CountCriteria(rubrics: Rubric[]): number {
    return rubrics.reduce(
      (accumulator, rubric) => {
        return accumulator + rubric.criteria.length;
      }, 0
    )
  }

  export interface Rubric {
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