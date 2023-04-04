import { getDirector, getMemberData } from '@hr/utils';
import { GS, SaDepartment } from '@lib/constants';
import { appendDataToSheet, copyInsides, createOrGetFolder, formatDate, getNamedValue, substituteVariables } from '@lib/functions';
import { File, Folder } from '@lib/models';
import { MemberModel } from '@models';
import { DocVariable, NamedRange, NamingConvention } from '@utils/constants';
import { memberToString } from '@utils/functions';

// TODO: add progress logs

export class Project {
  private defaultEdition = `${new Date().getFullYear()}.${new Date().getMonth() > 5 ? 2 : 1}`;

  start: Date;

  edition: string;

  manager: MemberModel;

  director: MemberModel;

  members: MemberModel[];

  constructor(public name: string, public department: SaDepartment) {
    this.edition = this.defaultEdition;
    this.start = new Date();
    this.members = [];
    this.director = getDirector(this.department);
  }

  /** Create project by reading data from the spreadsheet. */
  static spreadsheetFactory(): Project {
    return new this(getNamedValue(NamedRange.ProjectName), getNamedValue(NamedRange.ProjectDepartment) as SaDepartment)
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
      this.manager && !members.some(m => m.nUsp === this.manager.nUsp) && members.push(this.manager);
      this.members = members;
    }

    return this;
  }

  createFolder(): Folder {
    const departmentFolderIt = DriveApp.getFolderById(getNamedValue(NamedRange.DriveRoot)).getFoldersByName(
      this.department !== SaDepartment.Administrative
        ? `${NamingConvention.DepartmentFolderPrefix}${this.department}`
        : `${NamingConvention.AdministrativeFolderPrefix}${this.department}`,
    );

    if (!departmentFolderIt.hasNext()) {
      throw new ReferenceError("The Drive's root was not found.");
    }

    const departmentFolder = departmentFolderIt.next();
    const projectFolderIt = departmentFolder.getFoldersByName(this.name);
    const projectFolder = projectFolderIt.hasNext() ? projectFolderIt.next() : departmentFolder.createFolder(this.name);
    const targetDir = projectFolder.createFolder(this.edition);
    const templatesFolders = DriveApp.getFolderById(getNamedValue(NamedRange.ProjectCreationTemplatesFolderId)).getFolders();
    const { templateVariables } = this;

    // copy general templates to project folder
    while (templatesFolders.hasNext()) {
      const folder = templatesFolders.next();

      if (this.name.match(folder.getName())) {
        copyInsides(
          folder,
          targetDir,
          name => this.processTemplateName(name, templateVariables),
          file => this.processTemplateBody(file, templateVariables),
        );
      }
    }

    // copy project team spreadsheet template to project folder
    const membersSheetTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateId));
    const membersSheetFile = membersSheetTemplate.makeCopy(
      membersSheetTemplate.getName().replace(DocVariable.ProjectName, this.name).replace(DocVariable.ProjectEdition, this.edition),
      targetDir,
    );

    // write members list to project members sheet
    appendDataToSheet(
      this.members,
      SpreadsheetApp.open(membersSheetFile).getSheetByName(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateSheetName)),
      member => [member.name, member.nickname, member.email],
    );

    return targetDir;
  }

  /** Create the opening doc or just return it if it already exists (also force set last updated date). */
  createOrGetOpeningDoc(): File {
    const openingDocTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectOpeningDocId));
    const tmpDir = createOrGetFolder('.tmp', DriveApp.getFileById(GS.ss.getId()).getParents().next());
    const openingDocName = this.processTemplateName(openingDocTemplate.getName());
    const fileIterator = tmpDir.getFilesByName(openingDocName);

    if (fileIterator.hasNext()) {
      return fileIterator.next().setName(openingDocName);
    }

    const docFile = openingDocTemplate.makeCopy(this.processTemplateName(openingDocTemplate.getName()), tmpDir);

    this.processTemplateBody(docFile);

    return docFile;
  }

  private processTemplateName(name: string, templateVariables = this.templateVariables) {
    return Object.entries(templateVariables).reduce((title, [variable, value]) => title.replace(variable, value), name);
  }

  private processTemplateBody(file: File, templateVariables = this.templateVariables) {
    return substituteVariables(
      templateVariables,
      file.getMimeType() === MimeType.GOOGLE_SHEETS ? SpreadsheetApp.open(file) : DocumentApp.openById(file.getId()),
    );
  }

  private get templateVariables(): Record<DocVariable, string> {
    return {
      [DocVariable.MeetingType]: `${NamingConvention.ProjectMinutesPrefix}${this.name}`,
      [DocVariable.ProjectDepartment]: this.department || DocVariable.ProjectDepartment,
      [DocVariable.ProjectEdition]: this.edition,
      [DocVariable.ProjectManager]: (this.manager ? memberToString(this.manager) : '?') || DocVariable.ProjectManager,
      [DocVariable.ProjectDirector]: (this.director ? memberToString(this.director) : '?') || DocVariable.ProjectDirector,
      [DocVariable.ProjectName]: this.name,
      [DocVariable.ProjectStart]: formatDate(this.start),
      [DocVariable.ProjectMembers]: this.members.reduce((acc, cur) => `${acc}• ${memberToString(cur)}\n`, '') || DocVariable.ProjectMembers,
    };
  }
}
