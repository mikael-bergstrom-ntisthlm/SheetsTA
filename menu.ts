function SheetsTA() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu("SheetsTA")
    .addItem("Get list of active classrooms", "Menu.GetClassrooms")
    .addItem("Get roster from Classroom", "Menu.GetRoster")
    .addItem("Get list of assignments", "Menu.GetAssignments")
    .addItem("Get student submissions", "Menu.GetStudentSubmissions")
    .addItem("Sanitize Github URLs", "Menu.SanitizeGithubURLs")
    .addItem("Get document activity (weeks)", "Menu.GetDocActivityWeeks")
    .addItem("Get document activity (dates)", "Menu.GetDocActivityDates")
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
  
      const docUrl: string = String(row[0]);
      const userResourceName: string = String(row[1]);
  
      const dates = DocsTA.GetHistory(docUrl, userResourceName);
  
      if (dates.length == 0) return [];
  
      const weeks: Set<string> = new Set(
        dates.map(date => {
          return Utilities.formatDate(date, Session.getScriptTimeZone(), "w");
        })
      );
  
      return Array.from(weeks).sort();
    })
  }

  export function GetDocActivityDates() {
    SheetsUtilsTA.ProcessCurrentRange(
      2, "First with gdocs links, second with user IDs",
      row => {
  
      const docUrl: string = String(row[0]);
      const userResourceName: string = String(row[1]);
  
      const dates = DocsTA.GetHistory(docUrl, userResourceName);
  
      if (dates.length == 0) return [];
  
      const dateStrings: Set<string> = new Set(
        dates.map(date => {
          return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
        })
      );
  
      return Array.from(dateStrings).sort();
    })
  }
}




// Scopes: https://github.com/labnol/apps-script-starter/blob/master/scopes.md

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