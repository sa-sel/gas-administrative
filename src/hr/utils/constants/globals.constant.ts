import { getNamedValue } from '@sa-sel/lib';
import { NamedRange } from '@utils/constants';

/** The global spreadsheet namespace. */
export class HRGS {
  /** The whole spreadsheet. */
  static ss = SpreadsheetApp.openById(getNamedValue(NamedRange.SheetIdHR));

  /** The Google Sheet's user interface. */
  static get ssui(): GoogleAppsScript.Base.Ui {
    return SpreadsheetApp.getUi();
  }
}
