import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { createProject as hrSheetSaveProject } from '@hr/utils';
import { DialogTitle, GS, confirm, fetchData, institutionalEmails, toString } from '@lib';
import { Logger, SafeWrapper } from '@lib/classes';
import { MemberModel } from '@models';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';

const dialogBody = `
Você tem certeza que deseja continuar com essa ação? Ela é irreversível e vai:
  - Criar a pasta do projeto no Drive da SA-SEL;
  - Salvar o documento de abertura do projeto em sua pasta no Drive;
  - Salvar os membros selecionados na pasta do projeto e em seu documento de abertura;
  - Enviar o documento de abertura e o link da pasta do projeto por email para os membros do projeto.
`;

const actuallyCreateProject = (project: Project, logger: Logger) => {
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
};

export const createProject = () =>
  SafeWrapper.factory(createProject.name, institutionalEmails).wrap((logger: Logger): void => {
    const project = Project.spreadsheetFactory();

    if (!project.name || !project.edition || !project.manager || !project.department) {
      throw Error('Estão faltando informações do projeto a ser aberto. São necessário pelo menos nome, edição, gerente e área.');
    }

    logger.log(DialogTitle.InProgress, `Execução iniciada para projeto "${project.name}".`);
    confirm(
      {
        title: `Abertura de Projeto: ${project.name} (${project.edition})`,
        body: dialogBody,
      },
      () => actuallyCreateProject(project, logger),
      logger,
    );
  });
