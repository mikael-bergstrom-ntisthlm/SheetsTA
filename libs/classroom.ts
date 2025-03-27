/// <reference path="./github.ts" />
/// <reference path="./interfaces.ts" />
/// <reference path="./drive.ts" />

namespace ClassroomTA {
  /**
   * Creates a roster of students by combining student lists from 1+ Google Classrooms
   * @param config {Config} The config object to use; will contain info on what classrooms to get rosters from
   * @returns {string[][]} A two-dimensional array; an array of rows containing student data. First row is headers. Each row (inner array) will contain) columns: Classroom (name), Course ID, Name, Surname, Email, UserID.
   */
  export function GetRoster(config: Config): string[][] {
    let rosterValues: string[][] = [["Classroom", "CourseID", "Name", "Surname", "Email", "UserID"]];

    config.pairs.forEach(pair => {

      // "1" means second column, where CourseID is stored
      if (rosterValues.some(row => row[1] === pair.courseID)) return; // Skip if course's roster is already processed

      const classroomName = Classroom.Courses?.get(pair.courseID).name ?? "unnamed classroom";

      let nextPageToken: string = "";

      do { // For each page of results

        // -- GET ROSTER
        const roster = Classroom.Courses?.Students?.list(pair.courseID,
          { pageToken: nextPageToken }
        );

        if (roster?.students == undefined) {
          SpreadsheetApp.getUi().alert("No roster found");
          return;
        }

        // -- PROCESS ROSTER
        roster.students.forEach(student => {

          // Skip if student already exists in roster
          //  "5" is the column for "UserID"
          if (rosterValues.some(row => row[5] === student.profile?.id)) return;

          rosterValues.push(
            [
              classroomName,
              pair.courseID,
              student.profile?.name?.givenName ?? "",
              student.profile?.name?.familyName ?? "",
              student.profile?.emailAddress ?? "",
              student.profile?.id ?? ""
            ]
          )
        })

        nextPageToken = roster.nextPageToken ?? "";

      } while (nextPageToken != "");
    });

    return rosterValues;
  }

  /**
   * Creates a list of assignments by combining assignment lists from 1+ Google Classrooms
   * @param config {Config} The config object to use; will contain info on what classrooms to get assignments from
   * @returns {string[][]} A two-dimensional array; an array of rows containing assignment data. First row is headers. Each row (inner array) will contain) columns: Title, CourseID, CourseworkID
   */
  export function GetAssignments(config: Config): string[][] {
    let values: string[][] = [["Title", "CourseID", "CourseworkID"]];

    config.pairs.forEach(pair => {

      // -- GET ASSIGNMENTS
      const assignments = Classroom.Courses?.CourseWork?.list(pair.courseID);
      if (assignments?.courseWork == undefined) {
        SpreadsheetApp.getUi().alert("No assignments found");
        return;
      }

      // -- PROCESS ASSIGNMENTS
      assignments.courseWork.forEach(assignment => {
        values.push(
          [
            assignment.title ?? "",
            pair.courseID,
            assignment.id ?? "",
          ]
        )
      });
    });

    return values;
  }

  /**
   * Creates a list of active classrooms the current user has access to
   * @returns {string[][]} A two-dimensional array; an array of rows containing classroom data. First row is headers. Each row (inner array) will contain) columns: Course name, CourseID
   */
  export function GetClassrooms(): string[][] {
    let classroomValues: string[][] = [["Course name", "CourseID"]];
    let nextPageToken: string = "";

    do {

      // -- GET CLASSROOMS
      const classrooms = Classroom.Courses?.list(
        {
          courseStates: ["ACTIVE"],
          pageToken: nextPageToken
        }
      )

      if (classrooms?.courses == undefined) {
        SpreadsheetApp.getUi().alert("No classrooms found!");
        return [];
      }

      // -- PROCESS CLASSROOMS
      classrooms.courses.forEach(course => {
        classroomValues.push(
          [
            course.name ?? "",
            course.id ?? ""
          ]
        )
      });

    } while (nextPageToken != "");

    return classroomValues;
  }

  /**
   * Creates a list of student submission attachments by combining such attachments from 1+ Google Classroom assignments
   * @param config {Config} The config object to use; will contain info on what classrooms & assignments to get submissions from
   * @returns {string[][]} A two-dimensional array; an array of rows containing submission attachment data. First row is headers. Each row (inner array) will contain) columns: UserID, CourseID, CourseworkID (assignment ID), State (turned in, created etc), MIME, Submission URL.
   */
  export function GetStudentSubmissions(config: Config): string[][] {
    let submissionValues: string[][] = [["UserID", "CourseID", "CourseworkID", "State", "Type", "MIME", "Submission URL"]];

    config.pairs.forEach(pair => {

      let nextPageToken: string = "";

      do {
        // -- GET SUBMISSIONS
        const submissions = Classroom.Courses?.CourseWork?.StudentSubmissions?.list(pair.courseID, pair.courseworkID,
          { pageToken: nextPageToken }
        );

        if (submissions?.studentSubmissions == undefined) {
          SpreadsheetApp.getUi().alert("No submissions found");
          return;
        }

        // -- PROCESS SUBMISSIONS
        submissions.studentSubmissions.forEach(submission => {
          if (submission.assignmentSubmission?.attachments == undefined) return;

          // -- PROCESS ATTACHMENTS
          submission.assignmentSubmission?.attachments.forEach(attachment => {

            // Prepare data
            const attachmentUrl = attachment.driveFile?.alternateLink ?? attachment.link?.url ?? attachment.youTubeVideo?.alternateLink ?? "unknown url";
            let attachmentType = GetAttachmentType(attachment)
            let attachmentMimeType = DriveTA.GetFileMimeType(attachment);

            submissionValues.push([
              submission.userId ?? "",
              pair.courseID,
              pair.courseworkID,
              submission.state ?? "",
              attachmentType,
              attachmentMimeType,
              GithubTA.UrlSanitize(attachmentUrl)
            ])
          });
        });

        nextPageToken = submissions.nextPageToken ?? "";

      } while (nextPageToken != "");

    })

    return submissionValues;
  }

  // ----------------------------------------------------------------------------
  //  HELPERS

  // TODO: Generalize this, reuse in master document
  export function GetConfigFromRange(range: GoogleAppsScript.Spreadsheet.Range): Config {

    const config: Config = {
      pairs: []
    }

    let cellContents: string = String(range?.getValue());

    const pairs = cellContents.split(",");

    pairs.forEach(pair => {
      const pairSeparated = pair.split("/");
      config.pairs.push(
        {
          courseID: pairSeparated[0].trim(),
          courseworkID: pairSeparated.length > 1 ? pairSeparated[1].trim() : "",
          targetSheetName: "_SUBMISSIONS"
        }
      )
    });

    return config;
  }

  /**
   * Check the type of a Google Classroom assignment attachment
   * @param attachment {GoogleAppsScript.Classroom.Schema.Attachment} - The attachment to check
   * @returns {string} - A string describing the attachment's type
   */
  export function GetAttachmentType(attachment: GoogleAppsScript.Classroom.Schema.Attachment): string {

    if (attachment.driveFile != undefined) return "Drive file";
    if (attachment.link != undefined) return "Link";
    if (attachment.youTubeVideo != undefined) return "Youtube video";

    return "Unknown type";
  }
}