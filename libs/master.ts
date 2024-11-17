namespace MasterDocument {
  export function Setup() {
    let spreadsheet = SpreadsheetApp.getActive();
    let setupSheet = spreadsheet.getSheetByName("_SETUP");
    if (!setupSheet) {
      SpreadsheetApp.getUi().alert("No _SETUP sheet found");
      return;
    }

    // Gimmeh pairs
    let pairs = GetPairsFromSetupSheet(setupSheet);
    if (!pairs) return;

    // Rosterize
    UpdateSheet("_ROSTER", pairs, spreadsheet, ClassroomTA.GetRosterFromPairsTo);

    // Get student submissions
    UpdateSheet("_SUBMISSIONS", pairs, spreadsheet, ClassroomTA.GetStudentSubmissionsFromPairsTo);
  }

  export function UpdateSheet(
    sheetName: string,
    pairs: ClassroomTA.ClassroomIdentifierPair[],
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    action: (pairs: ClassroomTA.ClassroomIdentifierPair[], origo: GoogleAppsScript.Spreadsheet.Range) => void) {

    spreadsheet.toast("Updating " + sheetName);

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.setFrozenRows(1);
    }
    else {
      sheet.clear();
    }

    const origo = sheet.getRange(1, 1);

    action(pairs, origo);
  }

  export function GetPairsFromSetupSheet(setupSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    let pairValues = setupSheet?.getRange("A1:B").getValues();
    if (!pairValues) return;

    const pairs: ClassroomTA.ClassroomIdentifierPair[] = [];

    pairValues?.forEach(row => {
      if (row[0] == "" || row[1] == "") return;

      pairs.push({
        courseID: String(row[0]),
        courseworkID: String(row[1])
      });
    });

    return pairs;
  }
}