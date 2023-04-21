import { Student, toString } from '@lib';
import { hrSheets } from '../constants';

const parseRowToMember = (row: any[]): Student =>
  new Student({
    name: toString(row[0]),
    nickname: toString(row[1]) || undefined,
    nUsp: toString(row[2]),
    email: toString(row[5]) || undefined,
    phone: toString(row[3]).asPhoneNumber() || undefined,
  });

export const getMemberData = (nusp: string): Student => {
  const occurrences = hrSheets.mainData.createTextFinder(nusp).findAll();

  if (occurrences.length !== 1) {
    throw new RangeError(`Occurrences of nUSP '${nusp}' found: ${occurrences.length}.`);
  }

  const rowIndex = occurrences.pop().getRow();
  const row = hrSheets.mainData.getRange(rowIndex, 1, 1, hrSheets.mainData.getLastColumn()).getValues().flat();

  return parseRowToMember(row);
};
