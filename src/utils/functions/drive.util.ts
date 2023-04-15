import { File, Folder, GS, createOrGetFolder, getNamedValue } from '@lib';
import { NamedRange } from '@utils/constants';

export const getTmpFolder = (): Folder => createOrGetFolder('.tmp', DriveApp.getFileById(GS.ss.getId()).getParents().next());

export const getOpeningDocTemplate = (): File => DriveApp.getFileById(getNamedValue(NamedRange.ProjectOpeningDocId));
