import { StudentBasicModel } from '@lib/models';

export type MemberModel = StudentBasicModel & { email: string };
