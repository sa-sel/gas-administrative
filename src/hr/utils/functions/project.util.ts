import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { HRGS, hrSheets, NamedRange } from '@hr/utils/constants';
import { addColsToSheet, appendDataToSheet } from '@lib';

/** Create a new project with its members.*/
export const createProject = (name: string, members: ProjectMemberModel[]) => {
  const newColData = [[name, undefined]];

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
};
