/// <reference path="./classroom.ts" />
/// <reference path="./sheets.ts" />

namespace MasterDocumentTA {
  export function Setup(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {

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

    return ConfigTA.GetFromRange(setupSheet.getRange("A1:C"))
    // return GetConfigFromSetupSheet(setupSheet);
  }

  export function UpdateRoster(masterConfig: ConfigTA.Config, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const rosterOrigo = SheetsTA.CreateOrGetSheet("_ROSTER", spreadsheet, true).getRange(1, 1);
    const rosterValues = ClassroomTA.GetRoster(masterConfig);
    SheetsTA.InsertValuesAt(rosterValues, rosterOrigo);
  }

  export function UpdateSubmissions(masterConfig: ConfigTA.Config, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    
    // Split into multiple config based on target sheet
    const configs: Map<string, ConfigTA.Config> = ConfigSplitByTargetSheet(masterConfig);

    configs.forEach((config, targetSheet) => {

      // Get student submissions
      const submissionValues = ClassroomTA.GetStudentSubmissions(config);
      const submissionsOrigo = SheetsTA.CreateOrGetSheet(targetSheet, spreadsheet, true).getRange(1, 1);

      SheetsTA.InsertValuesAt(submissionValues, submissionsOrigo);
    });
  }

  function ConfigSplitByTargetSheet(masterConfig: ConfigTA.Config) {
    const configs: Map<string, ConfigTA.Config> = new Map();

    masterConfig.pairs.forEach(pair => {
      let key = pair.targetSheetName;

      if (key === "") key = "_SUBMISSIONS";

      // Key missing? Add it, with an empty config
      if (!configs.has(key)) configs.set(key, { pairs: [] });

      configs.get(key)?.pairs.push(pair);
    });

    return configs;
  }
}