import { getBoardOfDirectors, getDirector, getMemberData } from '@hr/utils';
import { SaDepartment } from '@lib/constants';
import { appendDataToSheet, copyInsides, exportToPdf, getNamedValue, substituteVariables } from '@lib/functions';
import { sendEmail } from '@lib/functions/email.util';
import { File, Folder } from '@lib/models';
import { MemberModel } from '@models';
import { GeneralVariable, NamedRange, ProjectVariable } from '@utils/constants';
import { getOpeningDocTemplate, getTmpFolder, memberToHtmlLi, memberToString } from '@utils/functions';

// TODO: how to use "@views/" here?
import emailBodyHtml from '../../../src/views/create-project.email.html';

// TODO: add progress logs

export class Project {
  private defaultEdition = `${new Date().getFullYear()}.${new Date().getMonth() > 5 ? 2 : 1}`;

  start: Date;
  edition: string;
  fullDepartmentName: string;
  manager: MemberModel;
  director: MemberModel;
  members: MemberModel[];
  folder: Folder;
  departmentFolder: Folder;
  openingDoc: File;

  constructor(public name: string, public department: SaDepartment) {
    this.edition = this.defaultEdition;
    this.start = new Date();
    this.members = [];
    this.director = getDirector(this.department);
    this.fullDepartmentName = this.department === SaDepartment.Administrative ? this.department : `Diretoria de ${this.department}`;

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

  setManager(manager?: MemberModel): Project {
    if (manager) {
      this.manager = manager;
    }

    return this;
  }

  setDirector(director?: MemberModel): Project {
    if (director) {
      this.director = director;
    }

    return this;
  }

  setEdition(edition?: string): Project {
    if (edition) {
      this.edition = edition.trim();
    }

    return this;
  }

  setMembers(members?: MemberModel[]): Project {
    if (members) {
      this.manager && !members.some(m => m.nUsp === this.manager.nUsp) && members.push(this.manager);
      this.director && !members.some(m => m.nUsp === this.director.nUsp) && members.push(this.director);
      this.members = members;
    }

    return this;
  }

  createFolder(): Folder {
    if (!this.departmentFolder) {
      throw new ReferenceError("The Drive's root or the department's folder was not found.");
    }

    const projectFolderIt = this.departmentFolder.getFoldersByName(this.name);
    const projectFolder = projectFolderIt.hasNext() ? projectFolderIt.next() : this.departmentFolder.createFolder(this.name);
    const templatesFolders = DriveApp.getFolderById(getNamedValue(NamedRange.ProjectCreationTemplatesFolderId)).getFolders();

    this.folder = projectFolder.createFolder(this.edition);

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

    // export opening doc PDF to project folder
    const openingDocIt = getTmpFolder().getFilesByName(this.processStringTemplate(getOpeningDocTemplate().getName(), templateVariables));

    if (openingDocIt.hasNext()) {
      const openingDocTmp = openingDocIt.next();

      substituteVariables(openingDocTmp, templateVariables);
      this.openingDoc = exportToPdf(openingDocTmp).moveTo(this.folder);
      openingDocTmp.setTrashed(true);
    }

    const board = getBoardOfDirectors();

    this.setupProjectControlSpreadsheet(
      {
        ...templateVariables,
        [GeneralVariable.MinutesTemplate]: getNamedValue(NamedRange.MinutesProjectTemplate),
      },
      board,
    );

    sendEmail({
      subject: `Abertura de Projeto - ${this.name} (${this.edition})`,
      target: [...this.members.map(({ email }) => email), ...board.map(({ email }) => email)],
      htmlBody: this.processStringTemplate(emailBodyHtml),
      attachments: this.openingDoc && [this.openingDoc],
    });

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

  private setupProjectControlSpreadsheet(templateVariables: Record<string, string>, boardOfDirectors: MemberModel[]): void {
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

  private processStringTemplate(str: string, templateVariables = this.templateVariables): string {
    return Object.entries(templateVariables).reduce((result, [variable, value]) => result.replaceAll(variable, value), str);
  }

  private get templateVariables(): Record<ProjectVariable, string> {
    return {
      [ProjectVariable.Department]: this.department || ProjectVariable.Department,
      [ProjectVariable.FullDepartment]: this.fullDepartmentName || ProjectVariable.FullDepartment,
      [ProjectVariable.Edition]: this.edition,
      [ProjectVariable.Manager]: this.manager ? memberToString(this.manager) : ProjectVariable.Manager,
      [ProjectVariable.Director]: this.director ? memberToString(this.director) : ProjectVariable.Director,
      [ProjectVariable.ManagerEmail]: this.manager?.email || ProjectVariable.ManagerEmail,
      [ProjectVariable.DirectorEmail]: this.director?.email || ProjectVariable.DirectorEmail,
      [ProjectVariable.Name]: this.name,
      [ProjectVariable.Start]: this.start.asDateString(),
      [ProjectVariable.NumMembers]: this.members.length.toString() || ProjectVariable.NumMembers,
      [ProjectVariable.Members]: this.members.reduce((acc, cur) => `${acc}• ${memberToString(cur)}\n`, '') || ProjectVariable.Members,
      [ProjectVariable.MembersHtmlList]: this.members.reduce((a, c) => `${a}${memberToHtmlLi(c)}\n`, '') || ProjectVariable.MembersHtmlList,
      [ProjectVariable.FolderUrl]: this.folder?.getUrl() || ProjectVariable.FolderUrl,
    };
  }
}
