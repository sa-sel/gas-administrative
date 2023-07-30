import { GS } from '@lib';
import { NamedRange } from '@utils/constants';
import { createProject } from './create-project.feature';
import { createMeetingMinutes } from './meeting-minutes.feature';
import { createProjectOpeningDoc } from './opening-doc.feature';

export const onOpen = () => {
  GS.ss.getRangeByName(NamedRange.ProjectData).clearContent();
  GS.ss.getRangeByName(NamedRange.ProjectMembers).clearContent();
  GS.ss.getRangeByName(NamedRange.MeetingAttendees).clearContent();

  GS.ui
    .createMenu('[Projetos]')
    .addItem('Criar Documento de Abertura', createProjectOpeningDoc.name)
    .addItem('Salvar Projeto (pasta, membros, etc)', createProject.name)
    .addToUi();

  GS.ui.createMenu('[Reuniões]').addItem('Criar Ata de Reunião Administrativa', createMeetingMinutes.name).addToUi();
};
