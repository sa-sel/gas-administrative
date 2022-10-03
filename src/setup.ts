import { GS } from '@hr/lib';
import { getNamedValue } from '@lib';
import { NamedRange } from '@utils/constants';

GS.ss = SpreadsheetApp.openById(getNamedValue(NamedRange.SheetIdHR));
