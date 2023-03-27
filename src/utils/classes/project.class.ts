import { getDirector } from '@hr/utils/functions';
import { SaDepartment } from '@lib/constants';
import { appendDataToSheet, copyInsides, formatDate, getNamedValue, substituteVariables } from '@lib/functions';
import { Folder } from '@lib/models';
import { MemberModel } from '@models';
import { DocVariable, NamedRange, NamingConvention } from '@utils/constants';

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
    const departmentFolderIt = DriveApp.getFolderById(NamedRange.DriveRoot).getFoldersByName(
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

    // copy general templates to project folder
    while (templatesFolders.hasNext()) {
      const folder = templatesFolders.next();

      if (this.name.match(folder.getName())) {
        copyInsides(
          folder,
          targetDir,
          name => name.replace(DocVariable.ProjectName, this.name),
          file =>
            substituteVariables(
              {
                [DocVariable.MeetingType]: `${NamingConvention.ProjectMinutesPrefix}${this.name}`,
                [DocVariable.ProjectDepartment]: this.department,
                [DocVariable.ProjectEdition]: this.edition,
                [DocVariable.ProjectManager]: this.manager ?? '?',
                [DocVariable.ProjectName]: this.name,
                [DocVariable.ProjectStart]: formatDate(this.start),
              },
              SpreadsheetApp.open(file) ?? DocumentApp.openById(file.getId()),
            ),
        );
      }
    }

    // copy project members spreadsheet template to project folder
    const membersSheetTemplate = DriveApp.getFileById(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateId));
    const membersSheetFile = membersSheetTemplate.makeCopy(
      membersSheetTemplate.getName().replace(DocVariable.ProjectName, this.name),
      targetDir,
    );

    // write members list to project members sheet
    appendDataToSheet(
      this.members,
      SpreadsheetApp.open(membersSheetFile).getSheetByName(getNamedValue(NamedRange.ProjectMembersSpreadsheetTemplateSheetName)),
      member => [member.name, member.email],
    );

    return targetDir;
  }
}
