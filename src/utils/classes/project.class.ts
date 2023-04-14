import { getDirector, getMemberData } from '@hr/utils';
import { SaDepartment } from '@lib/constants';
import { appendDataToSheet, copyInsides, exportToPdf, formatDate, getNamedValue, substituteVariables } from '@lib/functions';
import { File, Folder } from '@lib/models';
import { MemberModel } from '@models';
import { DocVariable, NamedRange, NamingConvention } from '@utils/constants';
import { getTmpFolder, memberToString } from '@utils/functions';

// TODO: add progress logs

export class Project {
  private defaultEdition = `${new Date().getFullYear()}.${new Date().getMonth() > 5 ? 2 : 1}`;

  start: Date;

  edition: string;

  manager: MemberModel;

  director: MemberModel;

  members: MemberModel[];

  folder: Folder;

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
      this.director && !members.some(m => m.nUsp === this.director.nUsp) && members.push(this.director);
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
    let openingDocPdf: File;
    const openingDocIt = getTmpFolder().getFilesByName(
      this.processStringTemplate(NamingConvention.OpeningDocTemplanteName, templateVariables),
    );

    if (openingDocIt.hasNext()) {
      const openingDoc = openingDocIt.next();

      substituteVariables(openingDoc, templateVariables);
      openingDocPdf = exportToPdf(openingDoc).moveTo(this.folder);
      openingDoc.setTrashed(true);
    }
    // copy project team spreadsheet template to project folder
    const membersSheetTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateId));
    const membersSheetFile = membersSheetTemplate.makeCopy(
      membersSheetTemplate.getName().replace(ProjectVariable.Name, this.name).replace(ProjectVariable.Edition, this.edition),
      this.folder,
    );

    // write members list to project members sheet
    appendDataToSheet(
      this.members,
      SpreadsheetApp.open(membersSheetFile).getSheetByName(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateSheetName)),
      member => [member.name, member.nickname, member.email],
    );

    return targetDir;

    return this.folder;
  }

  /** Create the opening doc or just return it if it already exists (also force set last updated date). */
  createOrGetOpeningDoc(): File {
    const openingDocTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectOpeningDocId));
    const tmpDir = getTmpFolder();
    const openingDocName = this.processStringTemplate(NamingConvention.OpeningDocTemplanteName);
    const fileIterator = tmpDir.getFilesByName(openingDocName);

    if (fileIterator.hasNext()) {
      return fileIterator.next().setName(openingDocName);
    }

    const openingDoc = openingDocTemplate.makeCopy(openingDocName, tmpDir);

    substituteVariables(openingDoc, this.templateVariables);

    return openingDoc;
  }

  private processStringTemplate(name: string, templateVariables = this.templateVariables): string {
    return Object.entries(templateVariables).reduce((title, [variable, value]) => title.replace(variable, value), name);
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
