import {
  DialogTitle,
  DiscordEmbed,
  DiscordWebhook,
  File,
  MeetingVariable,
  SafeWrapper,
  SheetLogger,
  Student,
  Transaction,
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
  SafeWrapper.factory(createMeetingMinutes.name, { allowedEmails: [...institutionalEmails, ...getAllMembers().map(m => m.email)] }).wrap(
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

      let minutesFile: File;
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

      new Transaction(logger)
        .step({
          forward: () => {
            minutesFile = project.createMeetingMinutes();
            minutesFile.setName(substituteVariablesInString(minutesFile.getName(), variables));
          },
          backward: () => minutesFile?.setTrashed(true),
        })
        .step({
          forward: () => substituteVariablesInFile(minutesFile, variables),
        })
        .run();

      const body = `Ata criada com sucesso:\n${minutesFile.getUrl()}`;

      alert({ title: DialogTitle.Success, body });
      logger.log(DialogTitle.Success, body);
      webhook.post({ embeds: buildMeetingDiscordEmbeds(project, now, attendees, clerk) });
    },
  );
