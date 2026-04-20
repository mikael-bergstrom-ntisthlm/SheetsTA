export namespace LibRubrics {

  export function GetSafeTagName(name: string) {
    return name.toLocaleLowerCase()
        .trim()
        .replace(/[^\w]/g, '')
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
    gradeTag: string;
    criteria: Criteria[];
    columnNumber: number;
  }

  export interface Criteria {
    name: string;
    tag: string;
    active: boolean;
    grade: string;
    studentPassed?: boolean;
    columnNumber: number;
  }
} 