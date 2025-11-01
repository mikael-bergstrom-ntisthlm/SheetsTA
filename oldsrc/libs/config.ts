

export namespace ConfigTA {

  export interface Config {
    gitFormat?: string,
    driveFormat?: string,
    pairs:
    {
      courseID: string,
      courseworkID: string,
      targetSheetName: string
    }[]
  }

  /**
   * Creates a Config object based on data from a Range
   * @param configRange {GoogleAppsScript.Spreadsheet.Range} The Range to get data from
   * @returns {Config|undefined} A config object containing Course ID / Assignment ID pairs & other config data
   */
  export function GetFromRange(configRange: GoogleAppsScript.Spreadsheet.Range): Config | undefined {

    // -- GET VALUES
    let configValues = configRange.getValues();

    if (configValues[0].length < 2) {
      Browser.msgBox("Selected range must be at least 2 columns wide");
      return;
    }

    const config: Config = {
      gitFormat: "",
      driveFormat: "",
      pairs: []
    }

    // -- PROCESS
    configValues?.forEach(row => {
      // Skip empty rows
      if (row[0] == "" || row[1] == "") return;

      // All IDs are 100% numbers so use that to identify course IDs
      if (!isNaN(parseInt(row[0]))) {
        config.pairs.push({
          courseID: String(row[0]),
          courseworkID: String(row[1]),
          targetSheetName: row.length > 2 ? String(row[2]) : "_"
        });
      }
      else if (row[0] == "drive") {
        config.driveFormat = row[1];
      }
    });

    return config;
  }

    /**
   * Split the pairs of a config into multiple configs based on their target sheet
   * @param {LibConfig.Config} masterConfig - 
   * @returns 
   */
  function ConfigSplitByTargetSheet(masterConfig: Config): Map<string, Config> {

    const configs: Map<string, Config> = new Map();

    masterConfig.pairs.forEach(pair => {
      // Use target sheet as key for map; "_SUBMISSIONS" if empty
      let key = pair.targetSheetName === "" ? "_SUBMISSIONS" : pair.targetSheetName;

      // Key missing? Add it, with an empty config
      if (!configs.has(key)) configs.set(key, { pairs: [] });

      // Add this pair to the config at the key
      configs.get(key)?.pairs.push(pair);
    });

    return configs;
  }
}