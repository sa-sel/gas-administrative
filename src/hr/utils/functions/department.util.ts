import { ProjectRole } from '@hr/models';
import { hrSheets } from '@hr/utils/constants';
import { manageDataInSheets, SaDepartment, SaDepartmentAbbreviations } from '@lib';
import { MemberModel } from '@models';
import { getMemberData } from './member.util';

export const getDirector = (department: SaDepartment): MemberModel | null => {
  let nusp: string;
  let director: MemberModel;

  // get nusps
  manageDataInSheets(SaDepartmentAbbreviations[department] ?? department, [hrSheets.projectMemberships], cell => {
    const col = cell.getColumn();
    const sheet = cell.getSheet();

    if (cell.getRow() === 1 && col > sheet.getFrozenColumns()) {
      const nFrozenRows = sheet.getFrozenRows();
      const departmentCol = sheet.getRange(nFrozenRows, col, sheet.getMaxRows() - nFrozenRows, 1);
      const directorCell = departmentCol.createTextFinder(ProjectRole.Director).findNext();

      if (directorCell) {
        const data: string[] = sheet.getRange(directorCell.getRow(), 3, 1, 1).getValues().flat();

        nusp = data[0];
      }
    }
  });

  // get member data
  try {
    director = getMemberData(nusp);
  } catch (err) {
    if (!(err instanceof RangeError)) {
      throw err;
    }
  }

  return director;
};

export const getBoardOfDirectors = (): MemberModel[] => {
  const directorNusps: Set<string> = new Set();

  // get nusps
  manageDataInSheets(ProjectRole.Director, [hrSheets.projectMemberships], cell => {
    const row = cell.getRow();
    const sheet = cell.getSheet();

    if (row > sheet.getFrozenRows() && cell.getColumn() > sheet.getFrozenColumns()) {
      const data: string[] = sheet.getRange(row, 3, 1, 1).getValues().flat();

      directorNusps.add(data[0]);
    }
  });

  // get member data
  return Array.from(directorNusps).reduce((board, nusp) => {
    try {
      board.push(getMemberData(nusp));
    } catch (err) {
      if (!(err instanceof RangeError)) {
        throw err;
      }
    }

    return board;
  }, []);
};
