/// <reference path="./libs/classroom.ts" />
/// <reference path="./libs/docs.ts" />
/// <reference path="./libs/github.ts" />
/// <reference path="./libs/sheets.ts" />
/// <reference path="./libs/utils.ts" />
/// <reference path="./libs/master.ts" />

function SheetsTA() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu("SheetsTA")
    .addItem("Get list of active classrooms", "SheetsTA.Menu.GetClassrooms")
    .addItem("Get roster from Classroom", "SheetsTA.Menu.GetRoster")
    .addItem("Get list of assignments", "SheetsTA.Menu.GetAssignments")
    .addItem("Get student submissions", "SheetsTA.Menu.GetStudentSubmissions")
    .addItem("Sanitize Github URLs", "SheetsTA.Menu.SanitizeGithubURLs")
    .addSeparator()
    .addItem("Get document activity (weeks)", "SheetsTA.Menu.GetDocActivityWeeks")
    .addItem("Get document activity (dates)", "SheetsTA.Menu.GetDocActivityDates")
    .addSeparator()
    .addItem("Get github repo activity (weeks)", "SheetsTA.Menu.GetGithubRepoActivityWeeks")
    .addItem("Get github repo activity (dates)", "SheetsTA.Menu.GetGithubRepoActivityDates")
    .addSeparator()
    .addItem("Setup entire document", "SheetsTA.MasterDocument.Setup")
    .addItem("Update document roster", "SheetsTA.Menu.UpdateRoster")
    .addItem("Update document submissions", "SheetsTA.Menu.UpdateSubmissions")
    .addToUi();
}

function SheetsTAInternal() {
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
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Master document")
      .addItem("Setup", "MasterDocument.Setup")
      .addItem("Update roster", "Menu.UpdateRoster")
      .addItem("Update submissions", "Menu.UpdateSubmissions")
      .addItem("Update Git activity page", "Menu.UpdateGitPage")
    )
    .addToUi();
}

export namespace Menu {
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
    SheetsUtilsTA.ProcessCurrentRange(row => GetDocActivity(row, "w"));
  }

  export function GetDocActivityDates() {
    SheetsUtilsTA.ProcessCurrentRange(row => GetDocActivity(row, "yyyy-MM-dd"));
  }

  export function GetGithubRepoActivityDates() {
    SheetsUtilsTA.ProcessCurrentRange(row => GetGithubRepoActivity(row, "yyyy-MM-dd"));
  }

  export function GetGithubRepoActivityWeeks() {
    SheetsUtilsTA.ProcessCurrentRange(row => GetGithubRepoActivity(row, "w"));
  }

  export function UpdateRoster() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    // Gimmeh pairs
    const config = MasterDocument.GetConfigFromSetupSheet(setupSheet);
    if (!config?.pairs) return;

    // Update roster
    const rosterOrigo = MasterDocument.CreateOrUpdateSheet("_ROSTER", spreadsheet);
    ClassroomTA.GetRosterFromPairsTo(config.pairs, rosterOrigo)
  }

  export function UpdateSubmissions() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    // Gimmeh pairs
    const config = MasterDocument.GetConfigFromSetupSheet(setupSheet);
    if (!config?.pairs) return;

    // Get student submissions
    const submissionsOrigo = MasterDocument.CreateOrUpdateSheet("_SUBMISSIONS", spreadsheet);
    ClassroomTA.GetStudentSubmissionsFromPairsTo(config.pairs, submissionsOrigo);
  }

  export function UpdateGitPage() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    let pairs = MasterDocument.GetConfigFromSetupSheet(setupSheet)?.pairs;
    if (!pairs) return;
  }


  function GetDocActivity(row: any[], format: string) {

    const dates = DocsTA.GetEditDates(
      String(row[0]), // DocURL
      row.length > 1 ? String(row[1]) : undefined // User ID
    );

    return Utils.GetUniqueDateStrings(dates, format);
  }

  function GetGithubRepoActivity(row: any[], format: string) {
    const repo = GithubTA.InterpretURL(String(row[0]))
    if (repo == undefined) return []

    const dates = GithubTA.GetCommitDates(
      repo,
      row.length > 1 ? String(row[1]) : undefined
    );

    return Utils.GetUniqueDateStrings(dates, format);
  }
}

// Scopes: https://github.com/labnol/apps-script-starter/blob/master/scopes.md

// TODO: Make MasterDocument.UpdateSheet return the content it just entered into the sheet? Also reference to A1 of that sheet?
// TODO: Same with GetStudentSubmissionsFromPairsTo and similar
// TODO: UpdateGitPage = update page, with a filtered version of submission data. Then run Get Activity.
// TODO: Auto-run GitPage update when Submissions page updates, if it exists (same for drive. Make generic?)

/* Implement:
- Make submenus
- Activity pages
  - _SETUP support: git/drive, dates/weeks
- Grading support
  - Generate grading page from current overview sheet
    - Rubrics
    - Checkboxes
    - Dropdown student names + id
  - Copy student's info from overview
  - Clear sheet
  - Copy sheet data back to overview
  - Generate overview sheet
    - Based on template
      - Extra info on each student
        - Submission filter/join columns (with formulas)
      - Source of rubrics: url
- File management
  - Naming files (Surname Name Assignment?)
*/