import { LibGSheets } from "../libs/sheets";
import { LibConfig } from "../libs/config";


export namespace PageMasterConfig {

  const helpText: string =
    `For assignments:
   Key = course ID
   Value = assignment ID
   Extra data = name of sheet where assignment submissions go
  `;

  const masterSpreadsheetName = "_SETUP";

  /**
   * Creates or empties a master config page
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet to create a master config _SETUP in
   * @returns 
   */
  export function CreateOrUpdateSetupSheet(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
  ) {

    // -- PREP
    const masterConfigSheet = LibGSheets.CreateOrGetSheet(masterSpreadsheetName, spreadsheet, true);
    if (!masterConfigSheet) return;

    // -- SETUP SHEET
    LibGSheets.ClearSheet(masterConfigSheet)

    // Add headers
    masterConfigSheet.setFrozenRows(1);
    masterConfigSheet.getRange(1, 1, 1, 4)
      .setValues([[
        "Key",
        "Value",
        "Extra data",
        "Comment"
      ]])

    // Add helpful(?) reminder box
    masterConfigSheet.getRange(2, 6, 5, 3)
      .merge().setVerticalAlignment("top")
      .setWrap(true)
      .setValue(helpText);
  }

  /**
   * Get the master config (as defined in a _SETUP sheet) from a spreadsheet document
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet document to get the master config of
   * @returns {LibConfig.Config | undefined} Either a config or, if there's no _SETUP, 'undefined'
   */
  export function GetMasterConfig(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): LibConfig.Config | undefined {

    const masterConfigSheet = LibGSheets.CreateOrGetSheet(masterSpreadsheetName, spreadsheet, true);
    if (!masterConfigSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    return LibConfig.GetFromRange(masterConfigSheet.getRange("A2:C"))
  }
}