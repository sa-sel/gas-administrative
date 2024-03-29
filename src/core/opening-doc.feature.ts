import { DialogTitle, DiscordEmbed, DiscordWebhook, GS, SafeWrapper, SheetLogger, alert, getNamedValue, institutionalEmails } from '@lib';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';

const buildProjectDiscordEmbeds = (project: Project): DiscordEmbed[] => {
  const fields: DiscordEmbed['fields'] = [
    { name: 'Nome', value: project.name, inline: true },
    { name: 'Edição', value: project.edition, inline: true },
  ];

  fields.pushIf(project.director || project.manager, { name: '', value: '' });
  fields.pushIf(project.director, { name: 'Direção', value: project.director.toString(), inline: true });
  fields.pushIf(project.manager, { name: 'Gerência', value: project.manager.toString(), inline: true });

  return [
    {
      title: 'Documento de Abertura',
      url: project.openingDoc.getUrl(),
      timestamp: project.start.toISOString(),
      fields,
      author: {
        name: project.fullDepartmentName,
        url: project.departmentFolder?.getUrl(),
      },
    },
  ];
};

export const createProjectOpeningDoc = () =>
  SafeWrapper.factory(createProjectOpeningDoc.name, { allowedEmails: institutionalEmails }).wrap((logger: SheetLogger): void => {
    const project = Project.spreadsheetFactory();

    if (!project.name || !project.edition || !project.department) {
      GS.ss.getRangeByName(NamedRange.ProjectData).activate();
      throw Error('Estão faltando informações do projeto a ser aberto. São necessário pelo menos nome, edição e diretoria.');
    }

    logger.log(DialogTitle.InProgress, `Execução iniciada para projeto "${project.name}".`);

    const doc = project.createOrGetOpeningDoc();
    const boardWebhook = new DiscordWebhook(getNamedValue(NamedRange.WebhookBoardOfDirectors));

    let body: string;
    let title: DialogTitle;

    // check if the doc was just created or already existed
    if (doc.getDateCreated().valueOf() === doc.getLastUpdated().valueOf()) {
      if (project.upsert()) {
        logger.log('Insert realizado!', `O projeto "${project.name}" foi salvo na lista de Projetos Existentes.`);
      }

      body = `Documento de abertura do projeto "${project.name}" criado com sucesso. Acesse o link:\n${doc.getUrl()}`;
      title = DialogTitle.Success;
    } else {
      body =
        `Documento de abertura do projeto "${project.name}" já havia sido criado, porém ainda não foi salvo. ` +
        `Acesse o link:\n${doc.getUrl()}`;
      title = DialogTitle.Aborted;
    }

    logger.log(title, body, false);
    alert({ title, body });
    boardWebhook.post({ embeds: buildProjectDiscordEmbeds(project) });
  });
