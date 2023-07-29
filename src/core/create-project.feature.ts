import { ProjectRole } from '@hr/models';
import { ProjectMember, createProject as hrSheetSaveProject } from '@hr/utils';
import { DialogTitle, DiscordEmbed, GS, confirm, fetchData, getNamedValue, institutionalEmails, toString } from '@lib';
import { DiscordWebhook, SafeWrapper, SheetLogger } from '@lib/classes';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';

const dialogBody = `
Você tem certeza que deseja continuar com essa ação? Ela é irreversível e vai:
  - Criar a pasta do projeto no Drive da SA-SEL;
  - Salvar o documento de abertura do projeto em sua pasta no Drive;
  - Salvar os membros selecionados na pasta do projeto e em seu documento de abertura;
  - Enviar o documento de abertura e o link da pasta do projeto por email para os membros do projeto e no Discord da SA-SEL.
`;

const buildProjectDiscordEmbeds = (project: Project): DiscordEmbed[] => {
  const fields: DiscordEmbed['fields'] = [
    { name: 'Nome', value: project.name, inline: true },
    { name: 'Edição', value: project.edition, inline: true },
  ];

  fields.pushIf(project.openingDoc, { name: 'Documento de Abertura', value: project.openingDoc?.getUrl() });
  fields.pushIf(project.director || project.manager, { name: '', value: '' });
  fields.pushIf(project.director, { name: 'Direção', value: project.director?.toString(), inline: true });
  fields.pushIf(project.manager, { name: 'Gerência', value: project.manager?.toString(), inline: true });
  fields.pushIf(project.members.length, {
    name: `Equipe (${project.members.length})`,
    value: project.members.toString(),
  });

  return [
    {
      title: 'Abertura de Projeto',
      url: project.folder.getUrl(),
      timestamp: project.start.toISOString(),
      fields,
      author: {
        name: project.fullDepartmentName,
        url: project.departmentFolder?.getUrl(),
      },
    },
  ];
};

const getMembers = (project: Project): ProjectMember[] =>
  fetchData(GS.ss.getRangeByName(NamedRange.ProjectMembers), {
    filter: row => [project.director?.nUsp, project.manager?.nUsp].includes(row[2]) || (row[2] && row[5]),
    map: row => {
      const member = new ProjectMember({
        name: toString(row[0]),
        nickname: toString(row[1]) || undefined,
        nUsp: toString(row[2]),
        phone: toString(row[3]).asPhoneNumber() || undefined,
        email: toString(row[4]) || undefined,
      });

      if (member.nUsp === project.director?.nUsp) {
        member.role = ProjectRole.Director;
      } else if (member.nUsp === project.manager?.nUsp) {
        member.role = ProjectRole.Manager;
      }

      return member;
    },
  });

const actuallyCreateProject = (project: Project, logger: SheetLogger) => {
  if (project.upsert()) {
    logger.log('Insert realizado!', `O projeto "${project.name}" foi salvo na lista de Projetos Existentes.`);
  }

  const members = getMembers(project);

  logger.log(DialogTitle.InProgress, `Projeto "${project.name}" possui ${members.length} membros.`);

  hrSheetSaveProject(project.toString(), members);
  logger.log(DialogTitle.Success, `Projeto ${project.name} salvo na planilha do RH.`);

  const dir = project.setMembers(members).createFolder();

  logger.log(DialogTitle.Success, `Pasta no Drive criada: ${dir.getName()} (${dir.getUrl()})`);

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
      `não esquece de criar o canal do projeto aqui no Discord: **#${discordChannel}**.`,
    embeds,
  });
};

export const createProject = () =>
  SafeWrapper.factory(createProject.name, { allowedEmails: institutionalEmails }).wrap((logger: SheetLogger): void => {
    const project = Project.spreadsheetFactory();

    if (!project.name || !project.edition || !project.manager || !project.department) {
      throw Error('Estão faltando informações do projeto a ser aberto. São necessário pelo menos nome, edição, gerente e área.');
    }

    logger.log(DialogTitle.InProgress, `Execução iniciada para projeto "${project.name}".`);
    confirm(
      {
        title: `Abertura de Projeto: ${project.toString()}`,
        body: dialogBody,
      },
      () => actuallyCreateProject(project, logger),
      logger,
    );
  });
