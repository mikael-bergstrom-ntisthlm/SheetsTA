import { LibGClassroom } from "../libs/classroom";
import { LibConfig } from "../libs/config";
import { LibGSheets } from "../libs/sheets";

export namespace PageSubmissions {

  /**
   * Add or update a sheet of student submissions to course assignments as defined in a config
   * @param {LibConfig.Config} config - The config to get course/assignment info from
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet document to add/update submissions to
   */
  export function Update(config: LibConfig.Config, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {

    // Make one config per target sheet
    const configs: Map<string, LibConfig.Config> = LibConfig.ConfigSplitByTargetSheet(config);

    configs.forEach((config, targetSheet) => {

      // Get student submissions
      const submissionValues = LibGClassroom.GetStudentSubmissions(config);
      const submissionsOrigo = LibGSheets.CreateOrGetSheet(targetSheet, spreadsheet, true).getRange(1, 1);

      // Insert submissions
      LibGSheets.InsertValuesAt(submissionValues, submissionsOrigo);
    });
  }
}