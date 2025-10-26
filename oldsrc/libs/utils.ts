namespace Utils {
  export function GetUniqueDateStrings(dates: Date[], format: string)
  {
    const dateStrings: Set<string> = new Set(
      dates.map(date => {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
      })
    );
    return Array.from(dateStrings).sort();
  }
}