import { LibRubrics } from "../libs/rubrics.js";


// This page should act as SSOT for rubrics
export namespace PageRubrics {

  const _RubricsSheetName = "_RUBRICS";

  const _ColRubricTitle = 1;
  const _ColCriteriaName = 2;
  const _ColCriteriaGrade = 3;
  const _ColCriteriaActive = 4;

  /**
   * Get an array of Rubrics and criteria from a rubrics setup sheet
   * @param rubricsSheet {GoogleAppsScript.Spreadsheet.Sheet}
   * @returns {LibRubrics.Rubric[]}
   */
  export function GetRubrics(rubricsSheet: GoogleAppsScript.Spreadsheet.Sheet): LibRubrics.Rubric[] {

    let rubricsData: string[][] = rubricsSheet.getRange(
      1, 1,
      rubricsSheet.getLastRow(),
      rubricsSheet.getLastColumn()
    ).getValues();

    let rubrics: LibRubrics.Rubric[] = [];
    let currentRubric: LibRubrics.Rubric | undefined = undefined;

    rubricsData.forEach(rubricDataRow => {

      // Check if rubric column contains a new rubric title
      if (rubricDataRow[_ColRubricTitle - 1] !== "") {
        currentRubric = {
          name: rubricDataRow[_ColRubricTitle - 1],
          criteria: [],
          gradeTag: LibRubrics.GetSafeTagName(rubricDataRow[_ColRubricTitle - 1]) + "grade"
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
            active: rubricDataRow[_ColCriteriaActive - 1] ? true : false
          }
        )
      }
    });

    return rubrics;
  }

  /**
   * Get the default rubrics sheet (defined by _RubricsSheetName) from a spreadsheet file
   * @param spreadsheet 
   * @returns 
   */
  export function GetDefaultRubricsSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet):
    GoogleAppsScript.Spreadsheet.Sheet | null {

    return spreadsheet.getSheetByName(_RubricsSheetName);
  }

}