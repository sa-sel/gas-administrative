import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { createProject as hrSheetSaveProject } from '@hr/utils/functions';
import { getMemberData } from '@hr/utils/functions/member.util';
import { DialogTitle, GS, SaDepartment, fetchData, getNamedValue, institutionalEmails, toString } from '@lib';
import { Logger, SafeWrapper } from '@lib/classes';
import { MemberModel } from '@models';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';

export const createProject = () =>
  SafeWrapper.factory(createProject.name, institutionalEmails).wrap((logger: Logger): void => {
    const project = new Project(getNamedValue(NamedRange.ProjectName), getNamedValue(NamedRange.ProjectDepartment) as SaDepartment)
      .setEdition(getNamedValue(NamedRange.ProjectEdition))
      .setManager(getMemberData(getNamedValue(NamedRange.ProjectManager).split(' - ')[1]));

    logger.log(DialogTitle.InProgress, `Execução iniciada para projeto "${project.name}".`);

    const members: (MemberModel & ProjectMemberModel)[] = fetchData(GS.ss.getRangeByName(NamedRange.ProjectMembers), {
      filter: row => [project.director?.nUsp, project.manager?.nUsp].includes(row[2]) || (row[2] && row[4]),
      map: row => {
        let role: ProjectRole;
        const member: MemberModel = {
          name: toString(row[0]),
          nickname: toString(row[1]),
          nUsp: toString(row[2]),
          email: toString(row[3]),
        };

        switch (member.nUsp) {
          case project.director?.nUsp: {
            role = ProjectRole.Director;
            break;
          }

          case project.manager?.nUsp: {
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
    });

    logger.log(DialogTitle.InProgress, `Projeto "${project.name}" possui ${members.length} membros.`);

    hrSheetSaveProject(`${project.name} (${project.edition})`, members);
    logger.log(`${DialogTitle.Success}`, `Projeto ${project.name} salvo na planilha do RH.`);

    const dir = project.setMembers(members).createFolder();

    logger.log(`${DialogTitle.Success}`, `Pasta no Drive criada: ${dir.getName()} (${dir.getUrl()})`);
  })();
