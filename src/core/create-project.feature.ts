import { createProject as hrSheetSaveProject } from '@hr/core';
import { GS as HRGS } from '@hr/lib';
import { ProjectMemberModel, ProjectRole } from '@hr/models';
import { SheetName as SheetNameHR, sheets } from '@hr/utils/constants';
import { getDirector } from '@hr/utils/functions';
import { fetchData, getNamedValue, GS, SaDepartment, toString } from '@lib';
import { MemberModel } from '@models';
import { Project } from '@utils/classes';
import { NamedRange } from '@utils/constants';

/** Override HR spreadsheet's sheet objects (previously created from active sheet). */
const overrideSheetsHR = (): void => {
  const hrSpreadsheet = SpreadsheetApp.openById(getNamedValue(NamedRange.SheetIdHR));

  HRGS.ss = hrSpreadsheet;
  sheets.caringMembers = hrSpreadsheet.getSheetByName(SheetNameHR.CaringMembers);
  sheets.caringProjects = hrSpreadsheet.getSheetByName(SheetNameHR.CaringProjects);
  sheets.dashboard = hrSpreadsheet.getSheetByName(SheetNameHR.Dashboard);
  sheets.mainData = hrSpreadsheet.getSheetByName(SheetNameHR.MainData);
  sheets.meetingAttendance = hrSpreadsheet.getSheetByName(SheetNameHR.MeetingAttendance);
  sheets.meetingAttendanceChart = hrSpreadsheet.getSheetByName(SheetNameHR.MeetingAttendanceChart);
  sheets.newMembers = hrSpreadsheet.getSheetByName(SheetNameHR.NewMembers);
  sheets.projectMemberships = hrSpreadsheet.getSheetByName(SheetNameHR.ProjectMemberships);
};

export const createProject = (): void => {
  overrideSheetsHR();

  const project = new Project(getNamedValue(NamedRange.ProjectName), getNamedValue(NamedRange.ProjectDepartment) as SaDepartment)
    .setEdition(getNamedValue(NamedRange.ProjectEdition))
    .setManager(getNamedValue(NamedRange.ProjectManager));
  const director = getDirector(project.department);
  const members: (MemberModel & ProjectMemberModel)[] = fetchData(GS.ss.getRangeByName(NamedRange.ProjectMembers), {
    map: row => {
      let role: ProjectRole;
      const member: MemberModel = { name: toString(row[0]), nickname: toString(row[1]), nUsp: toString(row[2]), email: toString(row[3]) };

      switch (member.nUsp) {
        case director?.nUsp: {
          role = ProjectRole.Director;
          break;
        }

        case project.manager: {
          role = ProjectRole.Manager;
          break;
        }

        default: {
          role = ProjectRole.Member;
          break;
        }
      }

      return { ...member, role };
    },
    filter: row => row[2] && row[4],
  });

  hrSheetSaveProject(project.name, members);
  project.setMembers(members).createFolder();
};
