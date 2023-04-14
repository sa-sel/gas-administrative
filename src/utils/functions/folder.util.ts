import { GS, createOrGetFolder } from '@lib';

export const getTmpFolder = () => createOrGetFolder('.tmp', DriveApp.getFileById(GS.ss.getId()).getParents().next());
