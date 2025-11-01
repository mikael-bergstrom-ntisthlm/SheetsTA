import { LibGClassroom } from "../libs/classroom";
import { LibConfig } from "../libs/config";
import { LibGSheets } from "../libs/sheets";

export namespace PageRoster {

  const rosterSheetName = "_ROSTER";

  
  /**
   * Add or update a sheet with roster information of one or more classrooms as defined in a config
   * @param {LibConfig.Config} config - The config to get roster info from
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet document to add roster sheet to
   */
  export function Update(config: LibConfig.Config, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {

    const rosterSheet = LibGSheets.CreateOrGetSheet(rosterSheetName, spreadsheet, true);
    if (!rosterSheet) return;

    const rosterValues = LibGClassroom.GetRoster(config);

    LibGSheets.InsertValuesAt(rosterValues, rosterSheet.getRange(1, 1));
  }
}