import { GS } from '@lib';
import { createProject } from './create-project.feature';

export const onOpen = () => {
  // GS.ss.getRangeByName(NamedRange.ProjectData).clearContent();
  // GS.ss.getRangeByName(NamedRange.ProjectMembers).clearContent();

  GS.ssui.createMenu('[Projetos]').addItem('Criar Novo Projeto', createProject.name).addToUi();
};
