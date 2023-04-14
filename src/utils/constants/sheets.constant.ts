import { GS } from '@lib/constants';

export const enum SheetName {
  Docs = 'Documentação',
  Logs = 'Logs',
  Minutes = 'Ata',
  ProjectDashboard = 'Controle de Projeto',
  ProjectDatabase = 'Projetos Existentes',
}

export const sheets = {
  docs: GS.ss.getSheetByName(SheetName.Docs),
  logs: GS.ss.getSheetByName(SheetName.Logs),
  minutes: GS.ss.getSheetByName(SheetName.Minutes),
  projectDashboard: GS.ss.getSheetByName(SheetName.ProjectDashboard),
  projectDatabase: GS.ss.getSheetByName(SheetName.ProjectDatabase),
};
