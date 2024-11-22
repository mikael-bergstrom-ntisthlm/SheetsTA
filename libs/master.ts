/// <reference path="./classroom.ts" />


namespace MasterDocument {
  export function Setup() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    // Gimmeh pairs
    let config = GetConfigFromSetupSheet(setupSheet);
    if (!config || !config.pairs) return;

    // Rosterize
    const rosterOrigo = CreateOrUpdateSheet("_ROSTER", spreadsheet);
    const rosterValues = ClassroomTA.GetRoster(config);
    SheetsTA.InsertValuesAt(rosterValues, rosterOrigo);

    // Get student submissions
    const submissionsOrigo = CreateOrUpdateSheet("_SUBMISSIONS", spreadsheet);
    const submissionValues = ClassroomTA.GetStudentSubmissions(config);
    SheetsTA.InsertValuesAt(submissionValues, submissionsOrigo);
  }

  export function CreateOrUpdateSheet(
    sheetName: string,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): GoogleAppsScript.Spreadsheet.Range {

    spreadsheet.toast("Updating " + sheetName);

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.setFrozenRows(1);
    }
    else {
      sheet.clear();
    }

    return sheet.getRange(1, 1);
  }

  export function GetConfigFromSetupSheet(setupSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    let pairValues = setupSheet?.getRange("A1:B").getValues();
    if (!pairValues) return;

    const config: Config = {
      gitFormat: "",
      driveFormat: "",
      pairs: []
    }

    pairValues?.forEach(row => {
      if (row[0] == "" || row[1] == "") return;

      // All IDs are 100% numbers
      if (!isNaN(parseFloat(row[0]))) {

        config.pairs.push({
          courseID: String(row[0]),
          courseworkID: String(row[1])
        });
      }
      else if (row[0] == "git") {
        config.gitFormat = row[1];
      }
      else if (row[0] == "drive") {
        config.driveFormat = row[1];
      }
    });

    return config;
  }
}