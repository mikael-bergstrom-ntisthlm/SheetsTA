export namespace LibRubrics {

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

  export interface Criteria {
    name: string;
    tag: string;
    active: boolean;
    grade: string;
    studentPassed?: boolean;
    columnNumber: number;
  }
} 