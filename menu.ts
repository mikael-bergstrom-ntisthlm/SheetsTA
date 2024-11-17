/// <reference path="./libs/classroom.ts" />
/// <reference path="./libs/docs.ts" />
/// <reference path="./libs/github.ts" />
/// <reference path="./libs/sheets.ts" />
/// <reference path="./libs/utils.ts" />
/// <reference path="./libs/master.ts" />



function SheetsTA() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu("SheetsTA")
    .addItem("Get list of active classrooms", "Menu.GetClassrooms")
    .addItem("Get roster from Classroom", "Menu.GetRoster")
    .addItem("Get list of assignments", "Menu.GetAssignments")
    .addItem("Get student submissions", "Menu.GetStudentSubmissions")
    .addItem("Sanitize Github URLs", "Menu.SanitizeGithubURLs")
    .addSeparator()
    .addItem("Get document activity (weeks)", "Menu.GetDocActivityWeeks")
    .addItem("Get document activity (dates)", "Menu.GetDocActivityDates")
    .addSeparator()
    .addItem("Get github repo activity (weeks)", "Menu.GetGithubRepoActivityWeeks")
    .addItem("Get github repo activity (dates)", "Menu.GetGithubRepoActivityDates")
    .addSeparator()
    .addItem("Setup entire document", "MasterDocument.Setup")
    .addItem("Update document roster", "Menu.UpdateRoster")
    .addItem("Update document submissions", "Menu.UpdateSubmissions")
    .addToUi();
}

namespace Menu {
  export function GetRoster() {

    let range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    let pairs = ClassroomTA.GetClassroomAndCourseworkIDPairs(range);
    let targetRangeStart = range.offset(range.getHeight(), 0, 1, 1);

    ClassroomTA.GetRosterFromPairsTo(pairs, targetRangeStart)
  }

  export function GetStudentSubmissions() {

    const range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    let pairs = ClassroomTA.GetClassroomAndCourseworkIDPairs(range);

    if (pairs.length < 1 || pairs[0].courseID == "" || pairs[0].courseworkID == "") {
      SpreadsheetApp.getUi().alert("Expected one or more course/assignment pair in selected cell");
      return;
    }

    let targetRangeStart = range.offset(range.getHeight(), 0, 1, 1);

    ClassroomTA.GetStudentSubmissionsFromPairsTo(pairs, targetRangeStart);

  }

  export function GetAssignments() {
    const range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    let pairs = ClassroomTA.GetClassroomAndCourseworkIDPairs(range);
    let targetRangeStart = range.offset(range.getHeight(), 0, 1, 1);

    ClassroomTA.GetAssignmentsFromPairsTo(pairs, targetRangeStart);
  }

  export function GetClassrooms() {

    let range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    ClassroomTA.GetClassroomsTo(range);
  }

  export function SanitizeGithubURLs() {
    let range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (range == undefined) return;
    let values = range.getValues();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        values[r][c] = GithubTA.UrlSanitize(values[r][c])
      }
    }
    range.setValues(values);
  }

  export function GetDocActivityWeeks() {
    SheetsUtilsTA.ProcessCurrentRange(
      2, "First with gdocs links, second with user IDs",
      row => {

        const dates = DocsTA.GetHistory(
          String(row[0]), // DocURL
          String(row[1]) // User ID
        );

        return Utils.GetUniqueDateStrings(dates, "w");
      })
  }

  export function GetDocActivityDates() {
    SheetsUtilsTA.ProcessCurrentRange(
      2, "First with gdocs links, second with user IDs to filter for",
      row => {

        const dates = DocsTA.GetHistory(
          String(row[0]), // DocURL
          String(row[1]) // User ID
        );

        return Utils.GetUniqueDateStrings(dates, "yyyy-MM-dd");
      })
  }

  export function GetGithubRepoActivityDates() {
    SheetsUtilsTA.ProcessCurrentRange(
      1, "With github links",
      row => {

        const repo = GithubTA.InterpretURL(String(row[0]))
        if (repo == undefined) return []

        const dates = GithubTA.GetCommitDates(repo);

        return Utils.GetUniqueDateStrings(dates, "yyyy-MM-dd");
      }
    )
  }

  export function GetGithubRepoActivityWeeks() {
    SheetsUtilsTA.ProcessCurrentRange(
      1, "With github links",
      row => {

        const repo = GithubTA.InterpretURL(String(row[0]))
        if (repo == undefined) return []

        const dates = GithubTA.GetCommitDates(repo);

        return Utils.GetUniqueDateStrings(dates, "w");
      }
    )
  }

  export function UpdateRoster() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    // Gimmeh pairs
    let pairs = MasterDocument.GetPairsFromSetupSheet(setupSheet);
    if (!pairs) return;

    // Rosterize
    MasterDocument.UpdateSheet("_ROSTER", pairs, spreadsheet, ClassroomTA.GetRosterFromPairsTo);
  }

  export function UpdateSubmissions() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    // Gimmeh pairs
    let pairs = MasterDocument.GetPairsFromSetupSheet(setupSheet);
    if (!pairs) return;

    // Get student submissions
    MasterDocument.UpdateSheet("_SUBMISSIONS", pairs, spreadsheet, ClassroomTA.GetStudentSubmissionsFromPairsTo);
  }
}



function Test() {


}

// Scopes: https://github.com/labnol/apps-script-starter/blob/master/scopes.md

// TODO: Make user ID / email filtering optional. Just one column? Great, just get everything then
// TODO: Support multiple ID / email filters

/* Implement:
- Mass-processing
  x Document activity (dates|weeks)
  x Github commits (dates|weeks)
- Full setup of document incl sheets based on pairs in _SETUP sheet
  - Get pairs
  - Setup roster
  - Setup submissions
- Github: Get direct link to Program.cs?
*/