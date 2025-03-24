/// <reference path="./classroom.ts" />
/// <reference path="./sheets.ts" />

function TestMaster() {
  MasterDocument.Setup();
}

namespace MasterDocument {
  export function Setup() {

    let spreadsheet = SpreadsheetApp.getActive();
    let masterConfig = GetMasterConfig(spreadsheet);
    if (!masterConfig || !masterConfig.pairs) return;

    // Rosterize
    UpdateRoster(masterConfig, spreadsheet);

    // Update submissions
    UpdateSubmissions(masterConfig, spreadsheet);
  }

  export function GetMasterConfig(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet)
  {
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    return GetConfigFromSetupSheet(setupSheet);
  }

  export function UpdateRoster(masterConfig: Config, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const rosterOrigo = SheetsTA.CreateOrGetSheet("_ROSTER", spreadsheet, true).getRange(1, 1);
    const rosterValues = ClassroomTA.GetRoster(masterConfig);
    SheetsTA.InsertValuesAt(rosterValues, rosterOrigo);
  }

  export function UpdateSubmissions(masterConfig: Config, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    
    // Split into multiple config based on target sheet
    const configs: Map<string, Config> = ConfigSplitByTargetSheet(masterConfig);

    configs.forEach((config, targetSheet) => {

      // Get student submissions
      const submissionValues = ClassroomTA.GetStudentSubmissions(config);
      const submissionsOrigo = SheetsTA.CreateOrGetSheet(targetSheet, spreadsheet, true).getRange(1, 1);

      SheetsTA.InsertValuesAt(submissionValues, submissionsOrigo);
    });
  }

  function ConfigSplitByTargetSheet(masterConfig: Config) {
    const configs: Map<string, Config> = new Map();

    masterConfig.pairs.forEach(pair => {
      let key = pair.targetSheetName;

      if (key === "") key = "_SUBMISSIONS";

      // Key missing? Add it, with an empty config
      if (!configs.has(key)) configs.set(key, { pairs: [] });

      configs.get(key)?.pairs.push(pair);
    });

    return configs;
  }

  export function GetConfigFromSetupSheet(setupSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    let pairValues = setupSheet?.getRange("A1:C").getValues();
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
          courseworkID: String(row[1]),
          targetSheetName: String(row[2])
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