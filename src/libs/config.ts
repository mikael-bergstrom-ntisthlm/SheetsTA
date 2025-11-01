export namespace LibConfig {

  /**
   * Creates a Config object based on data from a Range
   * @param configRange {GoogleAppsScript.Spreadsheet.Range} The Range to get data from
   * @returns {Config|undefined} A config object containing Course ID / Assignment ID pairs & other config data
   */
  export function GetFromRange(configRange: GoogleAppsScript.Spreadsheet.Range): Config | undefined {

    // -- GET VALUES
    let configValues = configRange.getValues();

    const config: Config = {
      pairs: []
    }

    // -- PROCESS
    configValues?.forEach(row => {
      // Skip empty rows
      if (row[0] == "") return;

      // All IDs are 100% numbers so use that to identify course IDs
      if (!isNaN(parseInt(row[0]))) {
        config.pairs.push({
          courseID: String(row[0]),
          courseworkID: row.length > 1 ? String(row[1]) : "",
          targetSheetName: row.length > 2 ? String(row[2]) : "_"
        });
      }
    });

    return config;
  }


  /**
   * Split the pairs of a config into multiple configs based on their target sheet
   * @param {LibConfig.Config} config - The config whose pairs to split
   * @returns {Map<string, LibConfig.Config>} A map of configs, with target sheets as keys
   */
  export function ConfigSplitByTargetSheet(config: LibConfig.Config): Map<string, LibConfig.Config> {

    const configs: Map<string, LibConfig.Config> = new Map();

    config.pairs.forEach(pair => {
      // Use target sheet as key for map; "_SUBMISSIONS" if empty
      let key = pair.targetSheetName === "" ? "_SUBMISSIONS" : pair.targetSheetName;

      // Key missing? Add it, with an empty config
      if (!configs.has(key)) configs.set(key, { pairs: [] });

      // Add this pair to the config at the key
      configs.get(key)?.pairs.push(pair);
    });

    return configs;
  }


  export interface Config {
    pairs:
    {
      courseID: string,
      courseworkID: string,
      targetSheetName: string // TODO: Document this
    }[]
  }

}