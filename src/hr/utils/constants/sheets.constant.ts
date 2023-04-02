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

export class hrSheets {
  static get caringMembers() {
    return HRGS.ss.getSheetByName(hrSheetName.CaringMembers);
  }

  static get caringProjects() {
    return HRGS.ss.getSheetByName(hrSheetName.CaringProjects);
  }

  static get dashboard() {
    return HRGS.ss.getSheetByName(hrSheetName.Dashboard);
  }

  static get mainData() {
    return HRGS.ss.getSheetByName(hrSheetName.MainData);
  }

  static get meetingAttendance() {
    return HRGS.ss.getSheetByName(hrSheetName.MeetingAttendance);
  }

  static get meetingAttendanceChart() {
    return HRGS.ss.getSheetByName(hrSheetName.MeetingAttendanceChart);
  }

  static get newMembers() {
    return HRGS.ss.getSheetByName(hrSheetName.NewMembers);
  }

  static get projectMemberships() {
    return HRGS.ss.getSheetByName(hrSheetName.ProjectMemberships);
  }
}
