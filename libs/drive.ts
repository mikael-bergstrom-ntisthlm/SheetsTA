namespace DriveTA {
  export function GetFileMimeType(attachment: GoogleAppsScript.Classroom.Schema.Attachment): string {
    if (attachment.driveFile == undefined
      || attachment.driveFile.id == undefined
    ) return "";

    let file = Drive.Files?.get(attachment.driveFile.id);
    if (!file?.mimeType) return "";

    return file.mimeType
  }
}