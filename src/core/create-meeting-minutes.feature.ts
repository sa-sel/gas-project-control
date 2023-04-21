import {
  DialogTitle,
  DiscordEmbed,
  DiscordWebhook,
  MeetingVariable,
  SafeWrapper,
  SheetLogger,
  Student,
  alert,
  getNamedValue,
  institutionalEmails,
  substituteVariablesInFile,
  substituteVariablesInString,
} from '@lib';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';
import { getAllMembers, getPresentMembers } from '@utils/functions';

const buildMeetingDiscordEmbeds = (project: Project, meetingStart: Date, attendees: Student[], clerk: string): DiscordEmbed[] => {
  const fields: DiscordEmbed['fields'] = [];

  fields.pushIf(project.manager, { name: 'Gerência', value: project.manager?.toString(), inline: true });
  fields.pushIf(clerk, { name: 'Redação', value: clerk, inline: true });
  fields.pushIf(attendees.length, { name: `Presentes (${attendees.length})`, value: attendees.toString() });

  return [
    {
      title: 'Reunião de Projeto',
      url: project.meetingMinutes.getUrl(),
      timestamp: meetingStart.toISOString(),
      fields,
      author: {
        name: project.toString(),
        url: project.folder.getUrl(),
      },
    },
  ];
};

export const createMeetingMinutes = () =>
  SafeWrapper.factory(createMeetingMinutes.name, [...institutionalEmails, ...getAllMembers().map(m => m.email)]).wrap(
    (logger: SheetLogger) => {
      const project = Project.spreadsheetFactory();
      const clerk = getNamedValue(NamedRange.MeetingClerk);
      const attendees = getPresentMembers();

      if (!project.name || !project.edition || !project.manager || !project.department || !clerk || !attendees.length) {
        throw Error(
          'Estão faltando informações do projeto a ser aberto. São necessário pelo menos ' +
            'nome, edição, gerente, área, redator(a) da ata e os membros presentes.',
        );
      }
      logger.log(DialogTitle.InProgress, `Execução iniciada para reunião com ${clerk} na redação.`);

      const minutesFile = project.createMeetingMinutes();
      const now = new Date();
      const variables: Record<MeetingVariable, string> = {
        [MeetingVariable.Clerk]: clerk,
        [MeetingVariable.Date]: now.asDateString(),
        [MeetingVariable.ReverseDate]: now.asReverseDateString(),
        [MeetingVariable.Start]: now.asTime(),
        [MeetingVariable.End]: MeetingVariable.End,
        [MeetingVariable.MeetingAttendees]: attendees.toBulletpoints(),
      };
      const webhook = new DiscordWebhook(getNamedValue(NamedRange.DiscordWebhook));

      minutesFile.setName(substituteVariablesInString(minutesFile.getName(), variables));
      substituteVariablesInFile(minutesFile, variables);

      const body = `Ata criada com sucesso:\n${minutesFile.getUrl()}`;

      logger.log(DialogTitle.Success, body);
      alert({ title: DialogTitle.Success, body });
      webhook.url && webhook.post({ embeds: buildMeetingDiscordEmbeds(project, now, attendees, clerk) });
    },
  );
