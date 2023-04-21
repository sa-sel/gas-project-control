import {
  DialogTitle,
  DiscordEmbed,
  DiscordWebhook,
  MeetingVariable,
  SafeWrapper,
  SheetLogger,
  Student,
  getNamedValue,
  institutionalEmails,
  substituteVariableInString,
  substituteVariablesInFile,
} from '@lib';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';
import { getAllMembers, getPresentMembers } from '@utils/functions';

const buildMeetingDiscordEmbeds = (project: Project, meetingStart: Date, attendees: Student[], clerk: string): DiscordEmbed[] => {
  const fields: DiscordEmbed['fields'] = [
    { name: 'Nome', value: project.name, inline: true },
    { name: 'Edição', value: project.edition, inline: true },
    { name: 'Ata', value: project.meetingMinutes.getUrl() },
  ];

  fields.pushIf(attendees.length, { name: `Presentes (${attendees.length})`, value: attendees.toString() });
  fields.pushIf(project.director || project.manager, { name: '', value: '' });
  fields.pushIf(project.manager, { name: 'Gerência', value: project.manager?.toString(), inline: true });
  fields.pushIf(clerk, { name: 'Redação', value: clerk, inline: true });

  return [
    {
      title: 'Reunião de Projeto',
      url: project.folder.getUrl(),
      timestamp: meetingStart.toISOString(),
      fields,
      author: {
        name: project.fullDepartmentName,
        url: project.departmentFolder?.getUrl(),
      },
    },
  ];
};

export const createMeetingMinutes = () =>
  SafeWrapper.factory(createMeetingMinutes.name, [...institutionalEmails, ...getAllMembers().map(m => m.email)]).wrap(
    (logger: SheetLogger) => {
      const project = Project.spreadsheetFactory();
      const clerk = getNamedValue(NamedRange.MeetingClerk);

      if (!project.name || !project.edition || !project.manager || !project.department || !clerk) {
        throw Error(
          'Estão faltando informações do projeto a ser aberto. São necessário pelo menos nome, edição, gerente, área e redator(a) da ata.',
        );
      }
      logger.log(DialogTitle.InProgress, `Execução iniciada para reunião com ${clerk} na redação.`);

      const attendees = getPresentMembers();
      const minutesFile = project.createMeetingMinutes();
      const now = new Date();
      const variables: Record<MeetingVariable, string> = {
        [MeetingVariable.Clerk]: clerk,
        [MeetingVariable.Date]: now.asDateString(),
        [MeetingVariable.ReverseDate]: now.asReverseDateString(),
        [MeetingVariable.Start]: now.asTime(),
        [MeetingVariable.End]: MeetingVariable.End,
        [MeetingVariable.MeetingAttendees]: attendees.toBulletpoints(),
        [MeetingVariable.Type]: 'Projeto',
      };
      const webhook = new DiscordWebhook(getNamedValue(NamedRange.DiscordWebhook));

      minutesFile.setName(substituteVariableInString(minutesFile.getName(), variables));
      substituteVariablesInFile(minutesFile, variables);

      logger.log(`${DialogTitle.Success}`, `Ata criada com sucesso:\n${minutesFile.getUrl()}`);
      webhook.url && webhook.post({ embeds: buildMeetingDiscordEmbeds(project, now, attendees, clerk) });
    },
  );
