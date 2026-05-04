/// <reference types="google-apps-script" />

import { LibGClassroom } from "./libs/classroom.js";
import { LibConfig } from "./libs/config.js";
import { LibGSheets } from "./libs/sheets.js";
import { LibGDocs } from "./libs/docs.js";
import { LibUtils } from "./libs/utils.js";
import { LibGithub } from "./libs/github.js";
import { PageMasterConfig } from "./pages/masterconfig.js";
import { PageRoster } from "./pages/roster.js";
import { PageSubmissions } from "./pages/submissions.js";
import { PageStudentGrading } from "./pages/studentgrading.js";
import { PageGradingOverview } from "./pages/gradingoverview.js";
import { PageRubrics } from "./pages/rubrics.js";
import { PageResponse } from "./pages/response.js";

function Setup() {
  let ui = SpreadsheetApp.getUi();

  const prefix: string = "SheetsTA2.";

  ui.createMenu("SheetsTA2")
    .addItem("Get list of active classrooms", `${prefix}GetClassrooms`)
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Selected course ID")
        .addItem("Get roster from Classroom", `${prefix}GetRoster`)
        .addItem("Get list of assignments", `${prefix}GetAssignments`)
        .addItem("Get student submissions", `${prefix}GetStudentSubmissions`)
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Activity tracking")
        .addItem("Get document activity (weeks)", `${prefix}GetDocActivityWeeks`)
        .addItem("Get document activity (dates)", `${prefix}GetDocActivityDates`)
        .addSeparator()
        .addItem("Get github repo activity (weeks)", `${prefix}GetGithubRepoActivityWeeks`)
        .addItem("Get github repo activity (dates)", `${prefix}GetGithubRepoActivityDates`)
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Master config")
        .addItem("Create master config", `${prefix}MasterConfigCreate`)
        .addItem("Update roster", `${prefix}UpdateRoster`)
        .addItem("Update submissions", `${prefix}UpdateSubmissions`)
        .addItem("Update all", `${prefix}UpdateAll`)
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Overview sheet")
        .addItem("Setup grading overview sheet", `${prefix}SetupGradingOverviewSheet`)
        .addItem("Update active criteria in overview sheet", `${prefix}UpdateGradingOverviewActiveFromTemplate`)
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Grading sheets")
        .addItem("Setup student grading sheet", `${prefix}SetupStudentGradingSheet`)
        .addItem("Clear student grading sheet", `${prefix}ClearStudentGradingSheet`)
        .addSeparator()
        .addItem("Transfer to grading overview & clear", `${prefix}TransferFromStudentGradingToOverview`)
        .addItem("Transfer from master grading sheet", `${prefix}TransferFromOverviewToStudentGrading`)
    ).addSubMenu(
      SpreadsheetApp.getUi().createMenu("Grading responses")
        .addItem("Setup response document template", `${prefix}SetupResponseTemplate`)
        .addItem("Generate/Update response for student", `${prefix}GenerateResponseDocForStudent`)
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("Utilities")
        .addItem("Sanitize Github URLs", "SheetsTA2.SanitizeGithubURLs")
    )

    .addToUi();
  Logger.log("Inited");
}

/* -----------------------------------------------------------------------------
  DIRECT MANIPULATION
------------------------------------------------------------------------------*/
//#region Direct manipulation

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

//#endregion

/* -----------------------------------------------------------------------------
  ACTIVITY TRACKING
------------------------------------------------------------------------------*/
//#region Activity tracking

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

//#endregion

/* -----------------------------------------------------------------------------
  MASTER CONFIG
------------------------------------------------------------------------------*/
//#region Master config

function MasterConfigCreate() {
  PageMasterConfig.CreateOrUpdateSetupSheet(SpreadsheetApp.getActive());
}

function UpdateRoster() {
  const spreadsheet = SpreadsheetApp.getActive();
  const config = PageMasterConfig.GetMasterConfig(spreadsheet)
  if (!config || !spreadsheet) return;
  PageRoster.Update(config, spreadsheet);
}

function UpdateSubmissions() {
  const spreadsheet = SpreadsheetApp.getActive();
  const config = PageMasterConfig.GetMasterConfig(spreadsheet)
  if (!config || !spreadsheet) return;
  PageSubmissions.Update(config, spreadsheet);
}

function UpdateAll() {
  const spreadsheet = SpreadsheetApp.getActive();
  const config = PageMasterConfig.GetMasterConfig(spreadsheet)
  if (!config || !spreadsheet) return;
  PageMasterConfig.UpdateAllPages(config, spreadsheet);
}

//#endregion

/* -----------------------------------------------------------------------------
  GRADING SHEETS
------------------------------------------------------------------------------*/
//#region Grading sheets

function SetupGradingOverviewSheet() {
  const spreadsheet = SpreadsheetApp.getActive();
  const config = PageMasterConfig.GetMasterConfig(spreadsheet)
  if (!config || !spreadsheet) return;

  PageGradingOverview.Setup.Setup(spreadsheet, config);
}

function UpdateGradingOverviewActiveFromTemplate() {
  const spreadsheet = SpreadsheetApp.getActive();
  PageGradingOverview.Setup.UpdateActiveCriteriaFromTemplate(spreadsheet)
}

function SetupStudentGradingSheet() {
  PageStudentGrading.Setup.Setup(SpreadsheetApp.getActive());
}

function ClearStudentGradingSheet() {
  const studentGradingSheet = PageStudentGrading.GetDefaultStudentGradingSheet(SpreadsheetApp.getActive());
  if (!studentGradingSheet) return;

  PageStudentGrading.ClearGrading(studentGradingSheet);
}

function TransferFromStudentGradingToOverview() {
  const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
  const studentGradingSheet = PageStudentGrading.GetDefaultStudentGradingSheet(spreadsheet);
  const gradingOverviewSheet = PageGradingOverview.GetDefaultGradingOverviewSheet(spreadsheet);
  const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);
  if (!studentGradingSheet || !gradingOverviewSheet || !rubricsSheet) return;

  const userId = PageStudentGrading.GetSelectedUserId(studentGradingSheet);
  if (userId === "") return;

  const rubrics = PageStudentGrading.GetStudentGradingData(
    rubricsSheet,
    studentGradingSheet
  )

  PageGradingOverview.InsertRubricData(userId, rubrics, gradingOverviewSheet);
  PageStudentGrading.ClearGrading(studentGradingSheet);
}

function TransferFromOverviewToStudentGrading() {
  const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
  const studentGradingSheet = PageStudentGrading.GetDefaultStudentGradingSheet(spreadsheet);
  const gradingOverviewSheet = PageGradingOverview.GetDefaultGradingOverviewSheet(spreadsheet);
  const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);
  if (!studentGradingSheet || !gradingOverviewSheet || !rubricsSheet) return;

  const userId = PageStudentGrading.GetSelectedUserId(studentGradingSheet);
  if (userId === "") return;

  let student = PageGradingOverview.GetStudentDataRubrics(
    userId,
    rubricsSheet,
    gradingOverviewSheet
  );
  if (!student) return;

  PageStudentGrading.InsertStudentDataRubrics(student, studentGradingSheet)
}

//#endregion

/* -----------------------------------------------------------------------------
  RESPONSE DOCS
------------------------------------------------------------------------------*/
//#region Response docs

function SetupResponseTemplate() {
  PageResponse.Setup(SpreadsheetApp.getActive());
}

function GenerateResponseDocForStudent() {
  const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
  const gradingOverviewSheet = PageGradingOverview.GetDefaultGradingOverviewSheet(spreadsheet);
  const rubricsSheet = PageRubrics.GetDefaultRubricsSheet(spreadsheet);
  if (!gradingOverviewSheet || !rubricsSheet) return;

  // PageGradingOverview.GetSelectedStudent(gradingOverviewSheet, rubricsSheet);
  const rowBlocks = LibGSheets.GetFullWidthBlocksOfSelection(gradingOverviewSheet);

  const targetFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();

  PageGradingOverview.GenerateResponseDocuments(rowBlocks, targetFolder, gradingOverviewSheet)

  // ResponsePage.GenerateResponseDocument(
  //   gradingOverviewSheet,
  //   "105003234631509491556"
  // )
}
//#endregion

/* -----------------------------------------------------------------------------
  UTILS
------------------------------------------------------------------------------*/
//#region Utilities

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

//#endregion

// -----------------------------------------------------------------------------
// TODO: Unify Setup naming & structure for all pages (Setup sub-namespace etc)
// TODO: Setup grading overview page based on roster, rubrics & additional config(?)
//         Extra columns
// TODO: See single user's results (incl. rubric matrix)
// TODO: Generate / Update individual student response sheets
// TODO: Centralize logging, toasts & alerts
// TODO: Variation: gyarte-stuff?
// TODO: Internationalization, at least sv/en via Session.getActiveUserLocale?


// REFERENCES
//  https://github.com/tomoyanakano/clasp-typescript-template