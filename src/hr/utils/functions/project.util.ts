import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { HRGS, NamedRange, hrSheets } from '@hr/utils/constants';
import { DialogTitle, addColsToSheet, appendDataToSheet } from '@lib';
import { SheetLogger } from '@lib/classes';

/** Create a new project with its members.*/
export const createProject = (name: string, members: ProjectMemberModel[]) => {
  const logger = new SheetLogger('createProject', HRGS.ss);
  const newColData = [[name, undefined]];

  logger.log(DialogTitle.InProgress, `Criação do projeto "${name}", disparada na planilha do administrativo.`, false);

  if (members.length) {
    const allMembersNusp = HRGS.ss.getRangeByName(NamedRange.AllSavedNusps).getValues().flat();
    const projectRoleByNusp: Record<string, ProjectRole> = {};
    const projetMembersNusp = new Set(
      members.map(member => {
        projectRoleByNusp[member.nUsp] = member.role;

        return member.nUsp;
      }),
    );

    allMembersNusp.forEach(nUsp => {
      newColData[0].push(projetMembersNusp.has(nUsp) ? projectRoleByNusp[nUsp] : undefined);
    });
  }

  addColsToSheet(newColData, hrSheets.projectMemberships);
  appendDataToSheet([{ projectName: name }], hrSheets.caringProjects, ({ projectName }) => [projectName]);

  logger.log(DialogTitle.Success, `Criação do projeto "${name}", disparada na planilha do administrativo.`, false);
};
