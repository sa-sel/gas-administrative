import { getBoardOfDirectors, getDirector, getMemberData } from '@hr/utils';
import { BaseProject, Student, Transaction } from '@lib';
import { DialogTitle, GS, SaDepartment } from '@lib/constants';
import {
  appendDataToSheet,
  copyInsides,
  exportToPdf,
  getNamedValue,
  substituteVariablesInFile,
  substituteVariablesInString,
} from '@lib/functions';
import { sendEmail } from '@lib/functions/email.util';
import { File, Folder } from '@lib/models';
import { GeneralVariable, NamedRange, sheets } from '@utils/constants';
import { getOpeningDocTemplate, getTmpFolder } from '@utils/functions';

// TODO: how to use "@views/" here?
import emailBodyHtml from '../../../src/views/create-project.email.html';

export class Project extends BaseProject {
  openingDoc: File;

  constructor(...args: ConstructorParameters<typeof BaseProject>) {
    super(...args);
    this.director = this.department ? getDirector(this.department) : null;

    const departmentFolderIt = DriveApp.getFolderById(getNamedValue(NamedRange.DriveRoot))?.getFoldersByName(this.fullDepartmentName);

    if (departmentFolderIt.hasNext()) {
      this.departmentFolder = departmentFolderIt.next();
    }
  }

  /** Create project by reading data from the spreadsheet. */
  static spreadsheetFactory(): Project {
    const manager = getNamedValue(NamedRange.ProjectManager)?.split(' - ')[1];

    const project = new this(
      getNamedValue(NamedRange.ProjectName).trim(),
      getNamedValue(NamedRange.ProjectDepartment).trim() as SaDepartment,
    ).setEdition(getNamedValue(NamedRange.ProjectEdition));

    manager && project.setManager(getMemberData(manager));

    return project;
  }

  // TODO:
  // - lidar com doc de Guia de Projeto
  // - atualizar enums e variáveis na planilha do RH
  createFolder(): Folder {
    if (!this.departmentFolder) {
      throw new ReferenceError("The Drive's root or the department's folder was not found.");
    }

    const projectFolderIt = this.departmentFolder.getFoldersByName(this.name);
    const projectFolder = projectFolderIt.hasNext() ? projectFolderIt.next() : this.departmentFolder.createFolder(this.name);
    const templatesFolders = DriveApp.getFolderById(getNamedValue(NamedRange.ProjectCreationTemplatesFolderId)).getFolders();

    this.folder = projectFolder.createFolder(this.edition);
    GS.ss.toast('Pasta do projeto criada.\nCopiando templates e documentos...', DialogTitle.InProgress);

    const { templateVariables } = this;

    // copy general templates to project folder
    while (templatesFolders.hasNext()) {
      const folder = templatesFolders.next();

      if (this.name.match(folder.getName())) {
        copyInsides(
          folder,
          this.folder,
          name => substituteVariablesInString(name, templateVariables),
          file => substituteVariablesInFile(file, templateVariables),
        );
      }
    }
    GS.ss.toast('Templates e documentos copiados para a pasta do projeto.', DialogTitle.InProgress);

    // export opening doc PDF to project folder
    const openingDocIt = getTmpFolder().getFilesByName(substituteVariablesInString(getOpeningDocTemplate().getName(), templateVariables));

    if (openingDocIt.hasNext()) {
      const openingDocEditable = openingDocIt.next();

      substituteVariablesInFile(openingDocEditable, templateVariables);
      this.openingDoc = exportToPdf(openingDocEditable).moveTo(this.folder);
      openingDocEditable.setTrashed(true);

      GS.ss.toast('PDF do documento de abertura exportado.', DialogTitle.InProgress);
    } else {
      GS.ss.toast('Não havia um documento de abertura para o projeto.', DialogTitle.InProgress);
    }

    const board = getBoardOfDirectors();
    const target = [...this.members.map(({ email }) => email), ...board.map(({ email }) => email)];

    this.setupProjectControlSpreadsheet(
      {
        ...templateVariables,
        [GeneralVariable.MinutesTemplate]: getNamedValue(NamedRange.MinutesProjectTemplate),
      },
      board,
    );
    GS.ss.toast('Planilha de controle de projeto criada.', DialogTitle.InProgress);

    sendEmail({
      subject: `Abertura de Projeto - ${this.name} (${this.edition})`,
      target,
      htmlBody: substituteVariablesInString(emailBodyHtml, templateVariables),
      attachments: this.openingDoc && [this.openingDoc],
    });
    GS.ss.toast(`Emails enviados para ${target.length} membros.`, DialogTitle.InProgress);

    return this.folder;
  }

  /** Create the opening doc or just return it if it already exists (also force set last updated date). */
  createOrGetOpeningDoc(): File {
    const { templateVariables } = this;
    const openingDocTemplate = getOpeningDocTemplate();
    const tmpDir = getTmpFolder();
    const openingDocName = substituteVariablesInString(openingDocTemplate.getName(), templateVariables);
    const fileIterator = tmpDir.getFilesByName(openingDocName);

    if (fileIterator.hasNext()) {
      return fileIterator.next().setName(openingDocName);
    }

    new Transaction()
      .step({
        forward: () => (this.openingDoc = openingDocTemplate.makeCopy(openingDocName, tmpDir)),
        backward: () => this.openingDoc?.setTrashed(true),
      })
      .step({
        forward: () => substituteVariablesInFile(this.openingDoc, templateVariables),
      })
      .run();

    return this.openingDoc;
  }

  /**
   * Save project name to the project name DB if it isn't there already.
   * @returns boolean if the project was inserted
   */
  upsert(): boolean {
    const sheet = sheets.projectDatabase;
    const exists = sheet.createTextFinder(this.name).findNext();

    if (!exists) {
      appendDataToSheet([[undefined, this.name, undefined]], sheet);

      const startRow = sheet.getFrozenRows() + 1;
      const endRow = sheet.getLastRow() - startRow + 1;

      sheets.projectDatabase.getRange(startRow, 1, endRow, sheet.getMaxColumns()).sort(2);

      return true;
    }

    return false;
  }

  private setupProjectControlSpreadsheet(templateVariables: Record<string, string>, boardOfDirectors: Student[]): void {
    if (!this.folder) {
      return;
    }

    let projectControlSheetFile: File;

    new Transaction()
      .step({
        // copy project control spreadsheet template to project folder
        forward: () => {
          const membersSheetTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateId));

          projectControlSheetFile = membersSheetTemplate.makeCopy(
            substituteVariablesInString(membersSheetTemplate.getName(), templateVariables),
            this.folder,
          );
        },
        backward: () => projectControlSheetFile?.setTrashed(true),
      })
      .step({
        // substitutes variables in project control spreadsheet
        forward: () => substituteVariablesInFile(projectControlSheetFile, templateVariables),
      })
      .step({
        forward: () => {
          const [membersSheet, boardOfDirectorsSheet] = SpreadsheetApp.open(projectControlSheetFile).getSheets();

          // write members list to project control sheet
          appendDataToSheet(this.members, membersSheet, m => [undefined, m.name, m.nickname, m.email, undefined, undefined]);
          // write directors list to project control sheet
          appendDataToSheet(boardOfDirectors, boardOfDirectorsSheet, m => [m.name, m.nickname, m.email]);
        },
      })
      .run();
  }
}
