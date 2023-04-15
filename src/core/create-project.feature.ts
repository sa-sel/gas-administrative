import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { createProject as hrSheetSaveProject } from '@hr/utils';
import { DialogTitle, DiscordEmbed, GS, confirm, fetchData, getNamedValue, institutionalEmails, toString } from '@lib';
import { DiscordWebhook, Logger, SafeWrapper } from '@lib/classes';
import { MemberModel } from '@models';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';
import { memberToString, upsertProject } from '@utils/functions';

const dialogBody = `
Você tem certeza que deseja continuar com essa ação? Ela é irreversível e vai:
  - Criar a pasta do projeto no Drive da SA-SEL;
  - Salvar o documento de abertura do projeto em sua pasta no Drive;
  - Salvar os membros selecionados na pasta do projeto e em seu documento de abertura;
  - Enviar o documento de abertura e o link da pasta do projeto por email para os membros do projeto e no Discord da SA-SEL.
`;

const buildProjectDiscordEmbeds = (project: Project): DiscordEmbed[] => [
  {
    title: 'Abertura de Projeto',
    url: project.folder.getUrl(),
    timestamp: project.start.toISOString(),

    fields: [
      { name: 'Nome', value: project.name, inline: true },
      { name: 'Edição', value: project.edition, inline: true },
      project.openingDoc && { name: 'Documento de Abertura', value: project.openingDoc.getUrl() },
      (project.director || project.manager) && { name: '', value: '' },
      project.director && { name: 'Direção', value: memberToString(project.director), inline: true },
      project.manager && { name: 'Gerência', value: memberToString(project.manager), inline: true },
      project.members.length && {
        name: `Membros (total ${project.members.length})`,
        value: project.members.map(memberToString).join(', '),
      },
    ],
    author: {
      name: project.fullDepartmentName,
      url: project.departmentFolder?.getUrl(),
    },
  },
];

const actuallyCreateProject = (project: Project, logger: Logger) => {
  if (upsertProject(project.name)) {
    logger.log('Insert realizado!', `O projeto "${project.name}" foi salvo na lista de Projetos Existentes.`);
  }

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

  const boardWebhook = new DiscordWebhook(getNamedValue(NamedRange.WebhookBoardOfDirectors));
  const generalWebhook = new DiscordWebhook(getNamedValue(NamedRange.WebhookGeneral));
  const embeds: DiscordEmbed[] = buildProjectDiscordEmbeds(project);
  const discordChannel = project.name
    .toLowerCase()
    .replaceAll(' ', '-')
    .removeAccents()
    .replace(/[^\w\d-]/g, '');

  generalWebhook.post({ embeds });
  boardWebhook.post({
    content:
      'Olá Diretoria:tm: , tudo bem?\n' +
      `Vocês acabaram de abrir o projeto **${project.name}** e aqui estão suas informações.\n\n` +
      `OBS: ${project.director?.nickname || project.director?.name || project.fullDepartmentName}, ` +
      `não esquece de criar o canal do projeto aqui no Discord: **${discordChannel}**.`,
    embeds,
  });
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
