import { ProjectRole } from '@hr/models';
import { HRGS, NamedRange, hrSheets } from '@hr/utils/constants';
import { DialogTitle, addColsToSheet, appendDataToSheet } from '@lib';
import { SheetLogger, Transaction } from '@lib/classes';
import { ProjectMember } from '../classes';

/** Create a new project with its members.*/
export const createProject = (name: string, members: ProjectMember[]) => {
  const logger = new SheetLogger('createProject', HRGS.ss.getSheetByName('Logs'));
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

  new Transaction(logger)
    .step({
      forward: () => {
        addColsToSheet(newColData, hrSheets.projectMemberships);
        logger.log(DialogTitle.InProgress, `Projeto adicionado para a planilha ${hrSheets.projectMemberships.getSheetName()}`, false);
      },
      backward: () => {
        const lastCol = hrSheets.projectMemberships.getMaxColumns();
        const lastColValue = hrSheets.projectMemberships.getSheetValues(0, lastCol, 1, 1)?.pop()?.pop();

        if (lastColValue === name) {
          hrSheets.projectMemberships.deleteColumn(lastCol);
        }
      },
    })
    .step({
      forward: () => {
        appendDataToSheet([{ projectName: name }], hrSheets.caringProjects, ({ projectName }) => [projectName]);
        logger.log(DialogTitle.InProgress, `Projeto adicionado para a planilha ${hrSheets.caringProjects.getSheetName()}`, false);
      },
      backward: () => {
        const lastRow = hrSheets.caringProjects.getMaxRows();
        const lastRowValue = hrSheets.caringProjects.getSheetValues(lastRow, 0, 1, 1)?.pop()?.pop();

        if (lastRowValue === name) {
          hrSheets.caringProjects.deleteRow(lastRow);
        }
      },
    })
    .run();

  logger.log(DialogTitle.Success, `Criação do projeto "${name}", disparada na planilha do administrativo.`, false);
};

export const rollbackCreateProject = (name: string) => {
  const lastCol = hrSheets.projectMemberships.getMaxColumns();
  const lastColValue = hrSheets.projectMemberships.getSheetValues(0, lastCol, 1, 1)?.pop()?.pop();

  if (lastColValue === name) {
    hrSheets.projectMemberships.deleteColumn(lastCol);
  }

  const lastRow = hrSheets.caringProjects.getMaxRows();
  const lastRowValue = hrSheets.caringProjects.getSheetValues(lastRow, 0, 1, 1)?.pop()?.pop();

  if (lastRowValue === name) {
    hrSheets.caringProjects.deleteRow(lastRow);
  }
};
