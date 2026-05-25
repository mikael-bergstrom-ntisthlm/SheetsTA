export namespace LibRubrics {

  /* -----------------------------------------------------------------------------
    HELPER FUNCTIONS
  ------------------------------------------------------------------------------*/
  //#region Interfaces

  /**
   * Returns a copy of a given name (string), with all non-words (a-z,A-Z,0-9,_) removed
   * @param name 
   * @returns 
   */
  export function GetSafeTagName(name: string) {
    return name.toLocaleLowerCase()
      .trim()
      .replace(/[^\w]/g, '')
  }

  /**
   * Counts the total number of criteria in a collection (array) of rubrics
   * @param rubrics 
   * @returns 
   */
  export function CountCriteria(rubrics: Rubric[]): number {
    return rubrics.reduce(
      (accumulator, rubric) => accumulator + rubric.criteria.length,
      0
    )
  }

  /**
   * Gets a combined array of all the criteria from a collection (array) of rubrics
   * @param rubrics 
   * @returns 
   */
  export function GetAllCriteria(rubrics: Rubric[]): Criteria[] {
    return rubrics.flatMap((rubric) => rubric.criteria);
  }

  /**
 * Returns the total number of columns or rows needed to fit a set of rubrics.
 * This a number equal to all the criteria + 2 per rubric (for grad & spacing)
 * @param rubrics 
 * @returns 
 */
  export function GetTotalSizeNeeded(rubrics: LibRubrics.Rubric[]): number {
    return LibRubrics.CountCriteria(rubrics) + rubrics.length * 2;
  }

  /* -----------------------------------------------------------------------------
    INTERFACES
  ------------------------------------------------------------------------------*/
  //#region Interfaces

  export interface Rubric {
    name: string;
    studentGrade?: string;
    gradeTag: string;
    criteria: Criteria[];
  }

  export interface Criteria {
    name: string;
    tag: string;
    active: boolean;
    grade: string;
    studentPassed?: boolean;
  }

  //#endregion
} 