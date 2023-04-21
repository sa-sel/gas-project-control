import {
  BaseProject,
  File,
  GS,
  SaDepartment,
  Student,
  createOrGetFolder,
  getNamedValue,
  substituteVariableInString,
  substituteVariablesInFile,
} from '@lib';
import { NamedRange } from '@utils/constants';
import { getAllMembers } from '@utils/functions';

export class Project extends BaseProject {
  meetingMinutes: File;

  /** Create project by reading data from the spreadsheet. */
  static spreadsheetFactory(): Project {
    return new this(getNamedValue(NamedRange.ProjectName).trim(), getNamedValue(NamedRange.ProjectDepartment).trim() as SaDepartment)
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

    this.meetingMinutes = minutesTemplate.makeCopy(substituteVariableInString(minutesTemplate.getName(), templateVariables), targetDir);
    substituteVariablesInFile(this.meetingMinutes, templateVariables);

    return this.meetingMinutes;
  }
}
