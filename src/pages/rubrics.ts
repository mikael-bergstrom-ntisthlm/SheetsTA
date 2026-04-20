import { LibRubrics } from "../libs/rubrics.js";


// This page should act as SSOT for rubrics
export namespace PageRubrics {

  const _RubricsSheetName = "_RUBRICS";

  const _ColRubricTitle = 1;
  const _ColCriteriaName = 2;
  const _ColCriteriaGrade = 3;
  const _ColCriteriaActive = 4;

  // TODO: Document this
  export function GetRubrics(rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet): LibRubrics.Rubric[] {

    let rubricsData: string[][] = rubricsSheet.getRange(
      1, 1,
      rubricsSheet.getLastRow(),
      rubricsSheet.getLastColumn()
    ).getValues();

    let rubrics: LibRubrics.Rubric[] = [];
    let currentRubric: LibRubrics.Rubric | undefined = undefined;
    let currentColNumber: number = 0;

    rubricsData.forEach(rubricDataRow => {

      // Check if rubric column contains a new rubric title
      if (rubricDataRow[_ColRubricTitle - 1] !== "") {
        currentColNumber += 1;
        currentRubric = {
          name: rubricDataRow[_ColRubricTitle - 1],
          criteria: [],
          columnNumber: currentColNumber,
          gradeTag: LibRubrics.GetSafeTagName(rubricDataRow[_ColRubricTitle - 1])
        }
        rubrics.push(currentRubric);
      }

      // Criteria
      if (rubricDataRow[_ColCriteriaName - 1] !== ""
        && rubricDataRow[_ColCriteriaGrade - 1] !== ""
        && currentRubric
      ) {
        currentRubric.criteria.push(
          {
            name: rubricDataRow[_ColCriteriaName - 1],
            tag: LibRubrics.GetSafeTagName(
              rubricDataRow[_ColCriteriaName - 1]
            ),
            grade: rubricDataRow[_ColCriteriaGrade - 1],
            columnNumber: currentColNumber,
            active: rubricDataRow[_ColCriteriaActive - 1] ? true : false
          }
        )
      }

      currentColNumber++;
    });

    return rubrics;
  }

  export function GetDefaultRubricsSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_RubricsSheetName);
  }

}