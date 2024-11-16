/// <reference path="./classroom.ts" />
/// <reference path="./docs.ts" />
/// <reference path="./github.ts" />
/// <reference path="./sheets.ts" />

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
    .addItem("Get github repo activity (dates)", "Menu.GetGithubRepoActivityDates")
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

        return GetUniqueDateStrings(dates, "w");
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

        return GetUniqueDateStrings(dates, "yyyy-MM-dd");
      })
  }

  export function GetGithubRepoActivityDates() {
    SheetsUtilsTA.ProcessCurrentRange(
      1, "With github links",
      row => {

        const repo = GithubTA.InterpretURL(String[row[0]])
        if (repo == undefined) return []
        
        const dates = GithubTA.GetCommitDates(repo);
       
        return GetUniqueDateStrings(dates, "yyyy-MM-dd");
      }
    )
  }
}

function GetUniqueDateStrings(dates: Date[], format: string)
{
  const dateStrings: Set<string> = new Set(
    dates.map(date => {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
    })
  );
  return Array.from(dateStrings).sort();
}

function Test()
{
  GithubTA.GetCommitDates({
    user: "mikael-bergstrom-ntisthlm",
    name: "GenericPlatformer/commits"
  });
}

// https://developers.google.com/apps-script/guides/services/external


// Scopes: https://github.com/labnol/apps-script-starter/blob/master/scopes.md

// TODO: Make user ID / email filtering optional. Just one column? Great, just get everything then
// TODO: Support multiple ID / email filters

/* Implement:
- Mass-processing
  x Document activity (dates|weeks)
  - Github commits (dates|weeks)
- Full setup of document incl sheets based on pairs in _SETUP sheet
  - Get pairs
  - Setup roster
  - Setup submissions
- Get github activity
  - return: list of dates
- Github: Get direct link to Program.cs?
*/