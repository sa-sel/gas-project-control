import { GS } from '@lib';
import { NamedRange } from '@utils/constants';
import { createMeetingMinutes } from './create-meeting-minutes.feature';

export const onOpen = () => {
  GS.ss.getRangeByName(NamedRange.MeetingClerk).clearContent();
  GS.ss.getRangeByName(NamedRange.ProjectMembers).uncheck();

  GS.ui.createMenu('[Reunião]').addItem('Criar Ata', createMeetingMinutes.name).addToUi();
};
