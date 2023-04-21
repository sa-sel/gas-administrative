import { getBoardOfDirectors, getDirector, getMemberData } from '@hr/utils';
import { BaseProject, Student } from '@lib';
import { DialogTitle, GS, SaDepartment } from '@lib/constants';
import { appendDataToSheet, copyInsides, exportToPdf, getNamedValue, substituteVariables } from '@lib/functions';
import { sendEmail } from '@lib/functions/email.util';
import { File, Folder } from '@lib/models';
import { GeneralVariable, NamedRange, sheets } from '@utils/constants';
import { getOpeningDocTemplate, getTmpFolder } from '@utils/functions';

// TODO: how to use "@views/" here?
import emailBodyHtml from '../../../src/views/create-project.email.html';

export class Project extends BaseProject {
  departmentFolder: Folder;
  openingDoc: File;

  constructor(...args: ConstructorParameters<typeof BaseProject>) {
    super(...args);
    this.director = getDirector(this.department);

    const departmentFolderIt = DriveApp.getFolderById(getNamedValue(NamedRange.DriveRoot))?.getFoldersByName(this.fullDepartmentName);

    if (departmentFolderIt.hasNext()) {
      this.departmentFolder = departmentFolderIt.next();
    }
  }

  /** Create project by reading data from the spreadsheet. */
  static spreadsheetFactory(): Project {
    return new this(getNamedValue(NamedRange.ProjectName).trim(), getNamedValue(NamedRange.ProjectDepartment).trim() as SaDepartment)
      .setEdition(getNamedValue(NamedRange.ProjectEdition))
      .setManager(getMemberData(getNamedValue(NamedRange.ProjectManager)?.split(' - ')[1]));
  }

  // TODO:
  // - lidar com os novos templates da Padrão
  // - atualizar enums e variáveis na planilha do RH
  // - scripts de ata de RD/RG (envio por email pra SA-SEL inteira + #general + #diretoria) -> salvar log na planilha do admin
  // - script planilha de controle de projeto (criação de ata) -> salvar log na planilha de controle de projeto
  // - script de projeto (envio por email pra projeto + diretoria + #diretoria + #projeto)
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
          name => this.processStringTemplate(name, templateVariables),
          file => substituteVariables(file, templateVariables),
        );
      }
    }
    GS.ss.toast('Templates e documentos copiados para a pasta do projeto.', DialogTitle.InProgress);

    // export opening doc PDF to project folder
    const openingDocIt = getTmpFolder().getFilesByName(this.processStringTemplate(getOpeningDocTemplate().getName(), templateVariables));

    if (openingDocIt.hasNext()) {
      const openingDocTmp = openingDocIt.next();

      substituteVariables(openingDocTmp, templateVariables);
      this.openingDoc = exportToPdf(openingDocTmp).moveTo(this.folder);
      openingDocTmp.setTrashed(true);

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
      htmlBody: this.processStringTemplate(emailBodyHtml),
      attachments: this.openingDoc && [this.openingDoc],
    });
    GS.ss.toast(`Emails enviados para ${target.length} membros.`, DialogTitle.InProgress);

    return this.folder;
  }

  /** Create the opening doc or just return it if it already exists (also force set last updated date). */
  createOrGetOpeningDoc(): File {
    const openingDocTemplate = getOpeningDocTemplate();
    const tmpDir = getTmpFolder();
    const openingDocName = this.processStringTemplate(openingDocTemplate.getName());
    const fileIterator = tmpDir.getFilesByName(openingDocName);

    if (fileIterator.hasNext()) {
      return fileIterator.next().setName(openingDocName);
    }

    this.openingDoc = openingDocTemplate.makeCopy(openingDocName, tmpDir);
    substituteVariables(this.openingDoc, this.templateVariables);

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

    // copy project control spreadsheet template to project folder
    const membersSheetTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateId));
    const projectControlSheetFile = membersSheetTemplate.makeCopy(this.processStringTemplate(membersSheetTemplate.getName()), this.folder);

    // substitutes variables in project control spreadsheet
    substituteVariables(projectControlSheetFile, templateVariables);

    const [membersSheet, boardOfDirectorsSheet] = SpreadsheetApp.open(projectControlSheetFile).getSheets();

    // write members list to project control sheet
    appendDataToSheet(this.members, membersSheet, m => [undefined, m.name, m.nickname, m.email, undefined, undefined]);
    // write directors list to project control sheet
    appendDataToSheet(boardOfDirectors, boardOfDirectorsSheet, m => [m.name, m.nickname, m.email]);
  }
}
