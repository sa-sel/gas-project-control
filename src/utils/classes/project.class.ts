import {
  BaseProject,
  File,
  GS,
  SaDepartment,
  Student,
  Transaction,
  createOrGetFolder,
  getNamedValue,
  substituteVariablesInFile,
  substituteVariablesInString,
} from '@lib';
import { NamedRange } from '@utils/constants';
import { getAllMembers } from '@utils/functions';

export class Project extends BaseProject {
  meetingMinutes: File;

  /** Create project by reading data from the spreadsheet. */
  static spreadsheetFactory(): Project {
    return new this(
      getNamedValue(NamedRange.ProjectName).trim(),
      getNamedValue(NamedRange.ProjectDepartment).trim().replace('Diretoria de', '') as SaDepartment,
    )
      .setEdition(getNamedValue(NamedRange.ProjectEdition))
      .setManager(Student.fromNameNicknameString(getNamedValue(NamedRange.ProjectManager), { nUsp: '' }))
      .setDirector(Student.fromNameNicknameString(getNamedValue(NamedRange.ProjectDirector), { nUsp: '' }))
      .setMembers(getAllMembers())
      .setFolder(DriveApp.getFileById(GS.ss.getId()).getParents().next());
  }

  createMeetingMinutes(): File {
    const { templateVariables } = this;
    const minutesTemplate = DriveApp.getFileById(getNamedValue(NamedRange.MinutesTemplate));
    const targetDir = createOrGetFolder('Atas', this.folder);

    new Transaction()
      .step({
        forward: () =>
          (this.meetingMinutes = minutesTemplate.makeCopy(
            substituteVariablesInString(minutesTemplate.getName(), templateVariables),
            targetDir,
          )),
        backward: () => this.meetingMinutes?.setTrashed(true),
      })
      .step({
        forward: () => substituteVariablesInFile(this.meetingMinutes, templateVariables),
      })
      .run();

    return this.meetingMinutes;
  }
}
