/// <reference path="./github.ts" />
/// <reference path="interfaces.ts" />


namespace ClassroomTA {
  export function GetRoster(config: Config) {
    let values: string[][] = [["Classroom", "CourseID", "Name", "Surname", "Email", "UserID"]];

    config.pairs.forEach(pair => {

      const classroomName = Classroom.Courses?.get(pair.courseID).name ?? "unnamed classroom";

      let nextPageToken: string = "";

      do {
        const roster = Classroom.Courses?.Students?.list(pair.courseID,
          {
            pageToken: nextPageToken
          }
        );
        if (roster?.students == undefined) {
          SpreadsheetApp.getUi().alert("No roster found");
          return;
        }

        roster.students.forEach(student => {
          values.push(
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

    return values;
  }

  export function GetAssignments(config: Config): string[][] {
    let values: string[][] = [["Title", "CourseID", "CourseworkID"]];

    config.pairs.forEach(pair => {

      const assignments = Classroom.Courses?.CourseWork?.list(pair.courseID);
      if (assignments?.courseWork == undefined) {
        SpreadsheetApp.getUi().alert("No assignments found");
        return;
      }

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

  export function GetClassroomsTo(targetRangeStart: GoogleAppsScript.Spreadsheet.Range) {
    let values: string[][] = [["Course name", "CourseID"]];
    let nextPageToken: string = "";

    do {
      const classrooms = Classroom.Courses?.list(
        {
          courseStates: ["ACTIVE"],
          pageToken: nextPageToken
        }
      )

      if (classrooms?.courses == undefined) {
        SpreadsheetApp.getUi().alert("No classrooms found!");
        return;
      }

      classrooms.courses.forEach(course => {
        values.push(
          [
            course.name ?? "",
            course.id ?? ""
          ]
        )
      });

    } while (nextPageToken != "");


    let targetRange = targetRangeStart?.offset(0, 0, values.length, values[0].length);

    targetRange?.setValues(values);
  }

  export function GetStudentSubmissions(config: Config): string[][] {
    let values: string[][] = [["UserID", "CourseID", "CourseworkID", "State", "Type", "Submission URL"]];

    config.pairs.forEach(pair => {

      let nextPageToken: string = "";

      do {
        const submissions = Classroom.Courses?.CourseWork?.StudentSubmissions?.list(pair.courseID, pair.courseworkID,
          { pageToken: nextPageToken }
        );

        if (submissions?.studentSubmissions == undefined) {
          SpreadsheetApp.getUi().alert("No submissions found");
          return;
        }

        submissions.studentSubmissions.forEach(submission => {
          if (submission.assignmentSubmission?.attachments == undefined) return;

          submission.assignmentSubmission?.attachments.forEach(attachment => {

            const attachmentUrl = attachment.driveFile?.alternateLink ?? attachment.link?.url ?? attachment.youTubeVideo?.alternateLink ?? "unknown url";
            let attachmentType = GetAttachmentType(attachment)

            values.push([
              submission.userId ?? "",
              pair.courseID,
              pair.courseworkID,
              submission.state ?? "",
              attachmentType,
              GithubTA.UrlSanitize(attachmentUrl)
            ])
          });
        });

        nextPageToken = submissions.nextPageToken ?? "";

      } while (nextPageToken != "");

    })

    return values;
  }

  // ----------------------------------------------------------------------------
  //  HELPERS

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
          courseworkID: pairSeparated.length > 1 ? pairSeparated[1].trim() : ""
        }
      )
    });

    return config;
  }

  export function GetAttachmentType(attachment: GoogleAppsScript.Classroom.Schema.Attachment) {
    return attachment.driveFile != undefined ? "Drive file" :
      attachment.link != undefined ? "Link" :
        attachment.youTubeVideo != undefined ? "Youtube video" :
          "Unknown type";
  }
}