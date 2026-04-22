import { LibRubrics } from "./rubrics.js";

export namespace LibStudents {
  /* -----------------------------------------------------------------------------
    INTERFACES
  ------------------------------------------------------------------------------*/
  //#region Interfaces

  export interface StudentData {
    id: string,
    name: string,
    surname: string,
    email: string,
    gradingData?: GradingData,
  }

  export interface GradingData {
    rubrics: LibRubrics.Rubric[],
    comment: string
  }

  //#endregion
}