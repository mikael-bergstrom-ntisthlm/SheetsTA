namespace DocsTA {
  export function GetEditDates(docUrl: string, userID?: string): Date[] {

    let editTimestamps: Date[] = [];
    const document = DocumentApp.openByUrl(docUrl);

    // Get all edits
    let result = DriveActivity.Activity?.query({
      "ancestorName": "items/" + document.getId(),
      "filter": "detail.action_detail_case:EDIT"
    });

    if (!result) return [];

    // Go through all edits
    result.activities?.forEach(activity => {
      if (!activity.actors) return;

      // Go through all editors
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