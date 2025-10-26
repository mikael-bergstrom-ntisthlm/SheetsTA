namespace DocsTA {
  /**
   * Get a list of Dates when the given Google Docs document was edited 
   * @param docUrl {string} The url of the document
   * @param userID {string} [undefined] If used, only changes made by that user will be included
   * @returns {Date[]} An array of dates
   */
  export function GetEditDates(docUrl: string, userID?: string): Date[] {

    // -- PREP
    let editTimestamps: Date[] = [];
    const document = DocumentApp.openByUrl(docUrl);

    // -- GET EDITS
    let result = DriveActivity.Activity?.query({
      "ancestorName": "items/" + document.getId(),
      "filter": "detail.action_detail_case:EDIT"
    });

    if (!result) return [];

    // -- PROCESS EDITS
    result.activities?.forEach(activity => {
      if (!activity.actors) return;

      // -- PROCESS EDIT'S EDITORS
      activity.actors.forEach(actor => {
        if (!actor.user?.knownUser?.personName
          || !activity.timestamp
        ) return;

        if (userID === undefined || actor.user.knownUser.personName === "people/" + userID)
        {
          editTimestamps.push(new Date(activity.timestamp));
        }
      })
    })

    return editTimestamps;
  }
}