import { toString } from '@lib';
import { MemberModel } from '@models';
import { hrSheets } from '../constants';

export const parseRowToMember = (row: any[]): MemberModel => ({
  name: toString(row[0]),
  nickname: toString(row[1]),
  nUsp: toString(row[2]),
  email: toString(row[5]),
});

export const getMemberData = (nusp: string): MemberModel => {
  const occurrences = hrSheets.mainData.createTextFinder(nusp).findAll();

  if (occurrences.length !== 1) {
    throw new RangeError(`Occurrences of nUSP '${nusp}' found: ${occurrences.length}.`);
  }

  const rowIndex = occurrences.pop().getRow();
  const row = hrSheets.mainData.getRange(rowIndex, 1, 1, hrSheets.mainData.getLastColumn()).getValues().flat();

  return parseRowToMember(row);
};
