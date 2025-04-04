/// <reference path="./libs/classroom.ts" />
/// <reference path="./libs/docs.ts" />
/// <reference path="./libs/github.ts" />
/// <reference path="./libs/sheets.ts" />
/// <reference path="./libs/utils.ts" />
/// <reference path="./libs/master.ts" />
/// <reference path="./libs/rubrics.ts" />
/// <reference path="./pages/studentgradingsheet.ts" />

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
      .addSeparator()
      .addItem("Setup student grading sheet", prefix + "Menu.SetupStudentGradingSheet")
      .addItem("Transfer to master grading sheet & clear", prefix + "Menu.TransferToMasterSheet")
      .addItem("Transfer from master grading sheet", prefix + "Menu.TransferFromMasterSheet")
      .addItem("Clear student grading sheet", prefix + "Menu.ClearStudentGradingSheet")
      .addToUi();
  }

  export function GetRoster() {

    let range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    const config = ConfigTA.GetFromRange(range);
    if (!config) return;

    let rosterOrigo = range.offset(range.getHeight(), 0, 1, 1);

    const values = ClassroomTA.GetRoster(config);
    SheetsTA.InsertValuesAt(values, rosterOrigo);
  }

  export function GetStudentSubmissions() {

    const range = SpreadsheetApp.getActiveSheet().getActiveRange();
    if (!range) return;

    let config = ConfigTA.GetFromRange(range);
    if (!config) return;

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

    let config = ConfigTA.GetFromRange(range);
    if (!config) return;

    let assignmentsSheetOrigo = range.offset(range.getHeight(), 0, 1, 1);

    const values = ClassroomTA.GetAssignments(config);
    SheetsTA.InsertValuesAt(values, assignmentsSheetOrigo);
  }

  export function GetClassrooms() {

    let classroomsOrigo = SpreadsheetApp
      .getActiveSheet()
      .getActiveRange();
    if (!classroomsOrigo) return;

    const values = ClassroomTA.GetClassrooms();
    SheetsTA.InsertValuesAt(values, classroomsOrigo);
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

    let masterConfig = MasterDocumentTA.GetMasterConfig(spreadsheet);
    if (!masterConfig || !masterConfig.pairs) return;

    // Rosterize
    MasterDocumentTA.UpdateRoster(masterConfig, spreadsheet);
  }

  export function UpdateSubmissions() {
    let spreadsheet = SpreadsheetApp.getActive();

    let masterConfig = MasterDocumentTA.GetMasterConfig(spreadsheet);
    if (!masterConfig || !masterConfig.pairs) return;

    // Update submissions
    MasterDocumentTA.UpdateSubmissions(masterConfig, spreadsheet);

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

  export function SetupStudentGradingSheet() {

    StudentGradingSheetTA.Setup.CreateOrUpdateStudentGradingSheet(SpreadsheetApp.getActive(), "Bedömning");
  }

  export function TransferToMasterSheet() {
    const masterGradingSheet = SpreadsheetApp.getActive().getSheetByName("Bedömning"); // TODO: make more general
    if (!masterGradingSheet) return;

    const studentGradingSheet = SpreadsheetApp.getActive().getSheetByName("_STUDENTGRADE");
    if (!studentGradingSheet) return;

    const userId = StudentGradingSheetTA.GetSelectedUserId(studentGradingSheet);
    if (userId === "") return null;

    StudentGradingSheetTA.TransferToMasterSheet(
      masterGradingSheet,
      studentGradingSheet,
      userId,
      true);
  }

  export function TransferFromMasterSheet() {
    const masterGradingSheet = SpreadsheetApp.getActive().getSheetByName("Bedömning"); // TODO: make more general
    if (!masterGradingSheet) return;

    const studentGradingSheet = SpreadsheetApp.getActive().getSheetByName("_STUDENTGRADE");
    if (!studentGradingSheet) return;

    const userId = StudentGradingSheetTA.GetSelectedUserId(studentGradingSheet);
    if (userId === "") return null;

    StudentGradingSheetTA.ImportFromMasterSheet(
      masterGradingSheet,
      studentGradingSheet,
      userId);
  }

  export function ClearStudentGradingSheet() {
    const studentGradingSheet = SpreadsheetApp.getActive().getSheetByName("_STUDENTGRADE");
    if (!studentGradingSheet) return;

    StudentGradingSheetTA.ClearGrading(studentGradingSheet);
  }
}

// Scopes: https://github.com/labnol/apps-script-starter/blob/master/scopes.md

/* Implement:
x MIME types of attachments
- Grading support
  x Generate grading page from current overview sheet
    x Rubrics
    x Checkboxes
    x Dropdown student names + id
    x Clear sheet
    x Copy sheet data back to overview from grading page
  x Remove unnecessary rows & cols from grading sheet
  x Copy student's info from overview to grading page
  x Clear student name when copying stuff to the master sheet
  x Add box for comment to student
  x Warn if overwriting?
  - Give Grade columns the right active/inactive bool
  ? Defaults for grades (+configurable?)

- Student response sheets (No more Autocrat!!!)
    - Generate template sheet
    - Generate documents in subfolder
      - with URL
      - Selectively
    - Update documents

- Generate master overview sheet
    - Based on template
      - Extra info on each student
        - Submission filter/join columns (with formulas)
      - Source of rubrics: url
  
    
- Make a sidebar with easy access shortcuts questionmark
- File management
  - Renaming files (Surname Name Assignment?)
- Moar git
  - Handle github access token
  - _GIT activity page
*/