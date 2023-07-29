import { DialogTitle, Spreadsheet, alert, getNamedValue } from '@lib';
import { NamedRange } from '@utils/constants';

/** The HR spreadsheet namespace. */
export class HRGS {
  private static ssCache: Spreadsheet;

  /** The whole spreadsheet. */
  static get ss(): Spreadsheet {
    if (this.ssCache === undefined) {
      try {
        this.ssCache = SpreadsheetApp.openById(getNamedValue(NamedRange.SheetIdHR));
      } catch (err) {
        alert({
          title: DialogTitle.Error,
          body: 'Você não tem permissão para abrir a planilha do RH, portanto as features dessa planilha não estão disponíveis.',
        });
        this.ssCache = null;
        throw err;
      }
    }

    return this.ssCache;
  }
}
