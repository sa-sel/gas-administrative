import { GS } from '@lib';
import { createProject } from './create-project.feature';
import { createProjectOpeningDoc } from './opening-doc.feature';
import { NamedRange } from '@utils/constants';

export const onOpen = () => {
  GS.ss.getRangeByName(NamedRange.ProjectData).clearContent();
  GS.ss.getRangeByName(NamedRange.ProjectMembers).clearContent();
  GS.ss.getRangeByName(NamedRange.MeetingAttendees).clearContent();

  GS.ui
    .createMenu('[Projetos]')
    .addItem('Criar Documento de Abertura', createProjectOpeningDoc.name)
    .addItem('Salvar Projeto (pasta, membros, etc)', createProject.name)
    .addToUi();
};
