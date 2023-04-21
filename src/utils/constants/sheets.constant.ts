import { GS } from '@lib/constants';

export const enum SheetName {
  Dashboard = 'Dashboard',
  BoardOfDirectors = 'Diretoria',
  Logs = 'Logs',
  Params = 'Parâmetros',
}

export const sheets = {
  dashboard: GS.ss.getSheetByName(SheetName.Dashboard),
  boardOfDirectors: GS.ss.getSheetByName(SheetName.BoardOfDirectors),
  params: GS.ss.getSheetByName(SheetName.Params),
  logs: GS.ss.getSheetByName(SheetName.Logs),
};
