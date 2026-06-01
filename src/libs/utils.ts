export namespace LibUtils {
  export function GetUniqueDateStrings(dates: Date[], format: string) {
    const dateStrings: Set<string> = new Set(
      dates.map(date => {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
      })
    );
    return Array.from(dateStrings).sort();
  }

  export function GetArrayOfFilteredObjectPropertyValues(
    object: Object, startsWithText: string): any[] {
    return Object
        .entries(object)
        .filter(entry => entry[0].startsWith(startsWithText))
        .map(entry => entry[1])
  }
}