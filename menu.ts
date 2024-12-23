/// <reference path="./libs/classroom.ts" />
/// <reference path="./libs/docs.ts" />
/// <reference path="./libs/github.ts" />
/// <reference path="./libs/sheets.ts" />
/// <reference path="./libs/utils.ts" />
/// <reference path="./libs/master.ts" />
/// <reference path="./libs/rubrics.ts" />


function SheetsTASetup() { Menu.Setup("SheetsTA."); }
function SheetsTAInternal() { Menu.Setup(""); }

export namespace Menu {
  export function Setup(prefix: string) {
    let ui = SpreadsheetApp.getUi();

    ui.createMenu("SheetsTA")
      .addItem("Get list of active classrooms", prefix + "Menu.GetClassrooms")
      .addItem("Get roster from Classroom", prefix + "Menu.GetRoster")
      .addItem("Get list of assignments", prefix + "Menu.GetAssignments")
      .addItem("Get student submissions", prefix + "Menu.GetStudentSubmissions")
      .addItem("Sanitize Github URLs", prefix + "Menu.SanitizeGithubURLs")
      .addSubMenu(SpreadsheetApp.getUi().createMenu("Activity tracking")
        .addItem("Get document activity (weeks)", prefix + "Menu.GetDocActivityWeeks")
        .addItem("Get document activity (dates)", prefix + "Menu.GetDocActivityDates")
        .addSeparator()
        .addItem("Get github repo activity (weeks)", prefix + "Menu.GetGithubRepoActivityWeeks")
        .addItem("Get github repo activity (dates)", prefix + "Menu.GetGithubRepoActivityDates")
      )
      .addSubMenu(SpreadsheetApp.getUi().createMenu("Master document")
        .addItem("Setup", prefix + "MasterDocument.Setup")
        .addItem("Update roster", prefix + "Menu.UpdateRoster")
        .addItem("Update submissions", prefix + "Menu.UpdateSubmissions")
      )
      .addToUi();
  }

  export function GetRoster() {

    let range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    const config = ClassroomTA.GetConfigFromRange(range);
    let rosterOrigo = range.offset(range.getHeight(), 0, 1, 1);

    const values = ClassroomTA.GetRoster(config);
    SheetsTA.InsertValuesAt(values, rosterOrigo);
  }

  export function GetStudentSubmissions() {

    const range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    let config = ClassroomTA.GetConfigFromRange(range);

    if (config.pairs.length < 1 || config.pairs[0].courseID == "" || config.pairs[0].courseworkID == "") {
      SpreadsheetApp.getUi().alert("Expected one or more course/assignment pair in selected cell");
      return;
    }

    let submissionsSheetOrigo = range.offset(range.getHeight(), 0, 1, 1);

    const values = ClassroomTA.GetStudentSubmissions(config);
    SheetsTA.InsertValuesAt(values, submissionsSheetOrigo);
  }

  export function GetAssignments() {
    const range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    let config = ClassroomTA.GetConfigFromRange(range);
    let assignmentsSheetOrigo = range.offset(range.getHeight(), 0, 1, 1);

    const values = ClassroomTA.GetAssignments(config);
    SheetsTA.InsertValuesAt(values, assignmentsSheetOrigo);
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
    SheetsTA.ProcessCurrentRange(row => GetDocActivity(row, "w"));
  }

  export function GetDocActivityDates() {
    SheetsTA.ProcessCurrentRange(row => GetDocActivity(row, "yyyy-MM-dd"));
  }

  export function GetGithubRepoActivityDates() {
    SheetsTA.ProcessCurrentRange(row => GetGithubRepoActivity(row, "yyyy-MM-dd"));
  }

  export function GetGithubRepoActivityWeeks() {
    SheetsTA.ProcessCurrentRange(row => GetGithubRepoActivity(row, "w"));
  }

  export function UpdateRoster() {
    let spreadsheet = SpreadsheetApp.getActive();

    let masterConfig = MasterDocument.GetMasterConfig(spreadsheet);
    if (!masterConfig || !masterConfig.pairs) return;

    // Rosterize
    MasterDocument.UpdateRoster(masterConfig, spreadsheet);
  }

  export function UpdateSubmissions() {
    let spreadsheet = SpreadsheetApp.getActive();

    let masterConfig = MasterDocument.GetMasterConfig(spreadsheet);
    if (!masterConfig || !masterConfig.pairs) return;

    // Update submissions
    MasterDocument.UpdateSubmissions(masterConfig, spreadsheet);

  }

  function GetDocActivity(row: any[], format: string) {

    const dates = DocsTA.GetEditDates(
      String(row[0]), // DocURL
      row.length > 1 ? String(row[1]) : undefined // User ID
    );

    return Utils.GetUniqueDateStrings(dates, format);
  }

  function GetGithubRepoActivity(row: any[], format: string): string[] {
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

// TODO: Fix structure for generating grading page

/* Implement:
x Make submenus
- Grading support
  - Generate grading page from current overview sheet
    x Rubrics
    x Checkboxes
    x Dropdown student names + id
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
- Moar git
  - Handle github access token
  - _GIT activity page
*/