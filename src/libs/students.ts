import { LibRubrics } from "./rubrics.js";

export namespace LibStudents {

  /**
   * Takes a set of row-data and inserts it into a Student object, using a tag
   * map to determine which of the row's columns maps to which criterion
   * The student needs to already have rubrics! And the row-data needs to be full-width
   * @param student {LibStudents.StudentData}
   * @param tagColNumbers {Map<string, number>}
   * @param studentDataValues {any[][]}
    */
  export function InsertRowDataIntoStudent(
    student: LibStudents.StudentData,
    tagColNumbers: Map<string, number>,
    studentDataValues: string[]
  ): void {

    if (student.gradingData === undefined) {
      return;
    }

    // Go through the rubrics
    student.gradingData.rubrics.forEach(rubric => {
      rubric.criteria.forEach(criterion => {

        // Find the column with a matching tag
        const colNumber = tagColNumbers.get(criterion.tag);
        if (colNumber === undefined) {
          Browser.msgBox(`No column found for criterion '${criterion.name}'`);
          return;
        }

        criterion.studentPassed =
          studentDataValues[colNumber] == "✔";
      });

      // Find column of rubric's overall grade
      const gradeCol = tagColNumbers.get(rubric.gradeTag);
      if (gradeCol === undefined) {
        Browser.msgBox(`No grade column found for rubric "${rubric.name}"`);
      } else {
        // Save rubric grade
        rubric.studentGrade = studentDataValues[gradeCol];
      }
    });

    // Get comment
    const colNumber = tagColNumbers.get("comment");
    if (colNumber === undefined) {
      Browser.msgBox("No column found for comment");
    }
    else {
      student.gradingData.comment = studentDataValues[colNumber];
    }
  }


  export function GetStudentsDataFromValues(
    sourceValues: any[][],
    setup: StudentColumnSetup
  ): LibStudents.StudentData[] {

    const studentsData: LibStudents.StudentData[] = [];

    sourceValues.forEach(row => {
      // Skip empties
      if (row[setup.colUserId - 1].length === 0) return;

      studentsData.push({
        id: row[setup.colUserId - 1],
        name: row[setup.colName - 1],
        surname: row[setup.colSurname - 1],
        email: row[setup.colEmail - 1]
      });
    });

    return studentsData;
  }

  /* -----------------------------------------------------------------------------
    INTERFACES
  ------------------------------------------------------------------------------*/
  //#region Interfaces

  export interface StudentColumnSetup {
    colName: number,
    colSurname: number,
    colEmail: number,
    colUserId: number,
  }

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