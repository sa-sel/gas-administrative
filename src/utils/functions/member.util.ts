import { MemberModel } from '@models';

export const memberToString = (member: MemberModel) => `${member.name} ${member.nickname ? '(' + member.nickname + ')' : ''}`.trim();
