interface InsertedValues {
  destinationOrigo: GoogleAppsScript.Spreadsheet.Range,
  values: string[][]
}

interface Config {
  gitFormat: string | undefined,
  driveFormat: string | undefined,
  pairs: ClassroomIdentifierPair[]
}

interface ClassroomIdentifierPair {
  courseID: string,
  courseworkID: string
}