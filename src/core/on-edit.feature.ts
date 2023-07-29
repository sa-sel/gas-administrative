import { GS, SheetsOnEditEvent } from '@lib';
import { NamedRange } from '@utils/constants';

let cachedMeetingTypeA1: string;

export const onEdit = (e: SheetsOnEditEvent) => {
  if (!cachedMeetingTypeA1) {
    cachedMeetingTypeA1 = GS.ss.getRangeByName(NamedRange.MeetingType).getA1Notation().slice(0, 2);
  }
  if (e.range.getA1Notation() === cachedMeetingTypeA1) {
    GS.ss.getRangeByName(NamedRange.MeetingAttendees).clearContent();
  }
};
