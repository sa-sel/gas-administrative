import { HRGS } from './globals.constant';

export const enum hrSheetName {
  CaringMembers = 'Acompanhamento - Membros',
  CaringProjects = 'Acompanhamento - Diretorias/Projetos',
  Dashboard = 'Home',
  MainData = 'Controle Geral',
  MeetingAttendance = 'Chamada RG',
  MeetingAttendanceChart = 'Gráfico (RGs)',
  NewMembers = 'Novos Membros',
  ProjectMemberships = 'Diretorias e Projetos',
}

export const hrSheets = {
  caringMembers: HRGS.ss.getSheetByName(hrSheetName.CaringMembers),
  caringProjects: HRGS.ss.getSheetByName(hrSheetName.CaringProjects),
  dashboard: HRGS.ss.getSheetByName(hrSheetName.Dashboard),
  mainData: HRGS.ss.getSheetByName(hrSheetName.MainData),
  meetingAttendance: HRGS.ss.getSheetByName(hrSheetName.MeetingAttendance),
  meetingAttendanceChart: HRGS.ss.getSheetByName(hrSheetName.MeetingAttendanceChart),
  newMembers: HRGS.ss.getSheetByName(hrSheetName.NewMembers),
  projectMemberships: HRGS.ss.getSheetByName(hrSheetName.ProjectMemberships),
};
