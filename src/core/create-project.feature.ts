import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { createProject as hrSheetSaveProject } from '@hr/utils/functions';
import { getMemberData } from '@hr/utils/functions/member.util';
import { fetchData, getNamedValue, GS, SaDepartment, toString } from '@lib';
import { MemberModel } from '@models';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';

export const createProject = (): void => {
  const project = new Project(getNamedValue(NamedRange.ProjectName), getNamedValue(NamedRange.ProjectDepartment) as SaDepartment)
    .setEdition(getNamedValue(NamedRange.ProjectEdition))
    .setManager(getNamedValue(NamedRange.ProjectManager));
  const director = getDirector(project.department);
  const members: (MemberModel & ProjectMemberModel)[] = fetchData(GS.ss.getRangeByName(NamedRange.ProjectMembers), {
    map: row => {
      let role: ProjectRole;
      const member: MemberModel = { name: toString(row[0]), nickname: toString(row[1]), nUsp: toString(row[2]), email: toString(row[3]) };

      switch (member.nUsp) {
        case director?.nUsp: {
          role = ProjectRole.Director;
          break;
        }

        case project.manager: {
          role = ProjectRole.Manager;
          break;
        }

        default: {
          role = ProjectRole.Member;
          break;
        }
      }

      return { ...member, role };
    },
    filter: row => row[2] && row[4],
  });

  hrSheetSaveProject(`${project.name} (${project.edition})`, members);
  project.setMembers(members).createFolder();
};
