export namespace LibRubrics {

  // TODO: Move to GradingOverviewSheet page file

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