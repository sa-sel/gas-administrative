import { DialogTitle, Logger, SafeWrapper, alert } from '@lib';
import { Project } from '@utils/classes';
import { upsertProject } from '@utils/functions';

export const createProjectOpeningDoc = () =>
  SafeWrapper.factory(createProjectOpeningDoc.name).wrap((logger: Logger): void => {
    const project = Project.spreadsheetFactory();

    if (!project.name || !project.edition) {
      throw Error('Estão faltando informações do projeto a ser aberto. São necessário pelo menos nome e edição.');
    }

    logger.log(DialogTitle.InProgress, `Execução iniciada para projeto "${project.name}".`);

    const doc = project.createOrGetOpeningDoc();

    // check if the doc was just created or already existed
    if (doc.getDateCreated().valueOf() === doc.getLastUpdated().valueOf()) {
      if (upsertProject(project.name)) {
        logger.log('Insert realizado!', `O projeto "${project.name}" foi salvo na lista de Projetos Existentes.`);
      }

      const body = `Documento de abertura do projeto "${project.name}" criado com sucesso. Acesse o link:\n${doc.getUrl()}`;

      logger.log(DialogTitle.Success, body, false);
      alert({ title: DialogTitle.Success, body });
    } else {
      const body = `Documento de abertura do projeto "${
        project.name
      }" já havia sido criado, porém ainda não foi salvo. Acesse o link:\n${doc.getUrl()}`;

      logger.log(DialogTitle.Aborted, body, false);
      alert({ title: DialogTitle.Success, body });
    }
  });
