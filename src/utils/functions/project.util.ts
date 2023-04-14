import { appendDataToSheet } from '@lib';
import { sheets } from '@utils/constants';

/**
 * Save project name to the project name DB if it isn't there already.
 * @returns boolean if the project was inserted
 */
export const upsertProject = (name: string): boolean => {
  const sheet = sheets.projectDatabase;
  const exists = sheet.createTextFinder(name).findNext();

  if (!exists) {
    appendDataToSheet([[undefined, name, undefined]], sheet);

    const startRow = sheet.getFrozenRows() + 1;
    const endRow = sheet.getLastRow() - startRow + 1;

    sheets.projectDatabase.getRange(startRow, 1, endRow, sheet.getMaxColumns()).sort(2);

    return true;
  }

  return false;
};
