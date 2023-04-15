import { ResourceUrl } from '@lib';
import { MemberModel } from '@models';

export const memberToString = (member: MemberModel) => `${member.name} ${member.nickname ? '(' + member.nickname + ')' : ''}`.trim();

export const memberToHtmlLi = (member: MemberModel) =>
  `<li>
    <a href="mailto:${member.email}" target="_blank">${memberToString(member)}</a>` +
  (member.phone
    ? ` — <a href="${ResourceUrl.WhatsAppApi.replace('{{phone}}', member.phone.replace(/\D/g, ''))}" target="_blank">${member.phone}</a>`
    : '') +
  `</li>`;
