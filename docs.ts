function Test()
{
  DocsTA.GetHistory(
    "https://docs.google.com/document/d/10gZsrYAC6aP6xtGmOeLR9pNgUmIsP6cFIGN3H22nulU/edit",
    "people/117946020979242483637"
  )
}

namespace DocsTA {
  export function GetHistory(docUrl: string, userID: string): Date[] {

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

        Logger.log(actor.user.knownUser.personName);
  
        if (actor.user.knownUser.personName === "people/" + userID) {
          editTimestamps.push(new Date(activity.timestamp));
        }
      })
    })
  
    return editTimestamps;
  }
}