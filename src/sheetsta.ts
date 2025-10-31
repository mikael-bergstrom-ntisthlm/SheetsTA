import { LibGClassroom } from "./libs/classroom";
import { LibConfig } from "./libs/config";
import { LibGSheets } from "./libs/sheets";
import { LibGDocs } from "./libs/docs";
import { LibUtils } from "./libs/utils";
import { LibGithub } from "./libs/github";

function Setup() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu("SheetsTA2")
    .addItem("Get list of active classrooms", "SheetsTA2.GetClassrooms")
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Selected course ID")
        .addItem("Get roster from Classroom", "SheetsTA2.GetRoster")
        .addItem("Get list of assignments", "SheetsTA2.GetAssignments")
        .addItem("Get student submissions", "SheetsTA2.GetStudentSubmissions")
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Activity tracking")
      .addItem("Get document activity (weeks)", "SheetsTA2.GetDocActivityWeeks")
      .addItem("Get document activity (dates)", "SheetsTA2.GetDocActivityDates")
      .addSeparator()
      .addItem("Get github repo activity (weeks)", "SheetsTA2.GetGithubRepoActivityWeeks")
      .addItem("Get github repo activity (dates)", "SheetsTA2.GetGithubRepoActivityDates")
    )
    // .addSubMenu(
    //   SpreadsheetApp.getUi().createMenu("Master config")
    //   // .addItem("Create master config", prefix + "MasterDocument.Create")
    //   // .addItem("Setup document", prefix + "MasterDocument.Setup")
    //   // .addItem("Update roster", prefix + "Menu.UpdateRoster")
    //   // .addItem("Update submissions", prefix + "Menu.UpdateSubmissions")
    // )
    // .addSubMenu(
    //   SpreadsheetApp.getUi().createMenu("Grading sheets")
    //   // .addItem("Setup student grading sheet", prefix + "Menu.SetupStudentGradingSheet")
    //   // .addItem("Transfer to master grading sheet & clear", prefix + "Menu.TransferToMasterSheet")
    //   // .addItem("Transfer from master grading sheet", prefix + "Menu.TransferFromMasterSheet")
    //   // .addItem("Clear student grading sheet", prefix + "Menu.ClearStudentGradingSheet")
    // )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Utilities")
      .addItem("Sanitize Github URLs", "SheetsTA2.SanitizeGithubURLs")
    )

    .addToUi();
  Logger.log("Inited");
}

// TODO: Move to its own LibDirectManipulation?
function GetClassrooms() {
  const classroomsOrigo = SpreadsheetApp
    .getActiveSheet()
    .getActiveRange();
  if (!classroomsOrigo) return;

  const values = LibGClassroom.GetClassrooms();
  LibGSheets.InsertValuesAt(values, classroomsOrigo);
}

function GetRoster() {

  let range = SpreadsheetApp.getActiveSheet().getActiveRange();
  if (!range) return;

  const config = LibConfig.GetFromRange(range);
  if (!config) return;

  let rosterOrigo = range.offset(range.getHeight(), 0, 1, 1);

  const values = LibGClassroom.GetRoster(config);
  LibGSheets.InsertValuesAt(values, rosterOrigo);
}

function GetAssignments() {
  const range = SpreadsheetApp.getActiveSheet().getActiveRange();
  if (!range) return;

  let config = LibConfig.GetFromRange(range);
  if (!config) return;

  let assignmentsSheetOrigo = range.offset(range.getHeight(), 0, 1, 1);

  const values = LibGClassroom.GetAssignments(config);
  LibGSheets.InsertValuesAt(values, assignmentsSheetOrigo);
}

function GetStudentSubmissions() {

  const range = SpreadsheetApp.getActiveSheet().getActiveRange();
  if (!range) return;

  let config = LibConfig.GetFromRange(range);
  if (!config) return;

  if (config.pairs.length < 1 || config.pairs[0].courseID == "" || config.pairs[0].courseworkID == "") {
    SpreadsheetApp.getUi().alert("Expected one or more course/assignment pair in selected cell");
    return;
  }

  let submissionsSheetOrigo = range.offset(range.getHeight(), 0, 1, 1);

  const values = LibGClassroom.GetStudentSubmissions(config);
  LibGSheets.InsertValuesAt(values, submissionsSheetOrigo);
}

/* -----------------------------------------------------------------------------
  ACTIVITY TRACKING
------------------------------------------------------------------------------*/

// TODO: Move to its own Activity lib?
function GetDocActivityWeeks() {
  LibGSheets.ProcessCurrentRange(row => GetDocActivity(row, "w"));
}

function GetDocActivityDates() {
  LibGSheets.ProcessCurrentRange(row => GetDocActivity(row, "yyyy-MM-dd"));
}

function GetGithubRepoActivityDates() {
  LibGSheets.ProcessCurrentRange(row => GetGithubRepoActivity(row, "yyyy-MM-dd"));
}

function GetGithubRepoActivityWeeks() {
  LibGSheets.ProcessCurrentRange(row => GetGithubRepoActivity(row, "w"));
}

function GetDocActivity(row: any[], format: string) {

  const dates = LibGDocs.GetEditDates(
    String(row[0]), // DocURL
    row.length > 1 ? String(row[1]) : undefined // User ID
  );

  return LibUtils.GetUniqueDateStrings(dates, format);
}

function GetGithubRepoActivity(row: any[], format: string): string[] {
  const repo = LibGithub.InterpretURL(String(row[0]))
  if (repo == undefined) return []

  const dates = LibGithub.GetCommitDates(
    repo,
    row.length > 1 ? String(row[1]) : undefined
  );

  return LibUtils.GetUniqueDateStrings(dates, format);
}



/* -----------------------------------------------------------------------------
  UTILS
------------------------------------------------------------------------------*/

function SanitizeGithubURLs() {
  let range = SpreadsheetApp.getActiveSheet().getActiveRange();
  if (range == undefined) return;
  let values = range.getValues();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      values[r][c] = LibGithub.UrlSanitize(values[r][c])
    }
  }
  range.setValues(values);
}