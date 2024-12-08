interface InsertedValues {
  destinationOrigo: GoogleAppsScript.Spreadsheet.Range,
  values: string[][]
}

interface Config {
  gitFormat?: string,
  driveFormat?: string,
  pairs:
  {
    courseID: string,
    courseworkID: string,
    targetSheetName: string
  }[]
}