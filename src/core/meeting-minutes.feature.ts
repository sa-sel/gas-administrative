import { getDirector } from '@hr/utils';
import {
  DialogTitle,
  DiscordEmbed,
  DiscordWebhook,
  File,
  GS,
  MeetingVariable,
  SaDepartment,
  SafeWrapper,
  SheetLogger,
  Student,
  Transaction,
  alert,
  areElementsInList,
  createOrGetFolder,
  fetchData,
  getNamedValue,
  institutionalEmails,
  substituteVariablesInFile,
  substituteVariablesInString,
  toString,
} from '@lib';
import { AdministrativeVariable, NamedRange } from '@utils/constants';

const buildDiscordEmbeds = (
  meetingType: string,
  meetingMinutes: File,
  meetingStart: Date,
  attendees: Student[],
  president: Student,
  vicePresident: Student,
  secretary: Student,
  clerk: string,
): DiscordEmbed[] => {
  const fields: DiscordEmbed['fields'] = [];
  const isClerkSecretary = clerk && clerk !== secretary?.toString();

  fields.pushIf(president, { name: SaDepartment.Presidency, value: president.toString(), inline: !isClerkSecretary });
  fields.pushIf(vicePresident, { name: SaDepartment.VicePresidency, value: vicePresident.toString(), inline: !isClerkSecretary });
  fields.pushIf(isClerkSecretary, { name: '', value: '', inline: false });
  fields.pushIf(secretary, { name: SaDepartment.Secretary, value: secretary.toString(), inline: !isClerkSecretary });
  fields.pushIf(clerk && clerk !== secretary?.toString(), { name: SaDepartment.Secretary, value: clerk, inline: !isClerkSecretary });
  fields.pushIf(attendees.length, { name: `Presentes (${attendees.length})`, value: attendees.toString(), inline: false });

  return [
    {
      title: meetingType,
      url: meetingMinutes.getUrl(),
      timestamp: meetingStart.toISOString(),
      fields,
      author: {
        name: 'SA-SEL',
        url: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
      },
    },
  ];
};

const getAttendees = () =>
  fetchData(GS.ss.getRangeByName(NamedRange.MeetingAttendees), {
    filter: row => row[2] && row[3],
    map: row =>
      new Student({
        name: toString(row[0]),
        nickname: toString(row[1]) || undefined,
        nUsp: toString(row[2]),
      }),
  });

const createMinutesFile = (
  meetingType: string,
  meetingStart: Date,
  attendees: Student[],
  president: Student,
  vicePresident: Student,
  secretary: Student,
  clerk: string,
): File => {
  let meetingMinutes: File;
  const minutesTemplate = DriveApp.getFileById(getNamedValue(NamedRange.MinutesAdminTemplate));
  const minutesDir = createOrGetFolder('Atas', DriveApp.getFileById(GS.ss.getId()).getParents().next());
  const targetDir = createOrGetFolder(meetingStart.getFullYear().toString(), minutesDir);

  const templateVariables: Record<MeetingVariable | AdministrativeVariable, string> = {
    [MeetingVariable.Clerk]: clerk,
    [MeetingVariable.Date]: meetingStart.asDateString(),
    [MeetingVariable.ReverseDate]: meetingStart.asReverseDateString(),
    [MeetingVariable.ReverseDateWithoutYear]: meetingStart.asReverseDateStringWithoutYear(),
    [MeetingVariable.Start]: meetingStart.asTime(),
    [MeetingVariable.End]: MeetingVariable.End,
    [MeetingVariable.MeetingAttendees]: attendees.toBulletpoints(),
    [MeetingVariable.MeetingType]: meetingType,
    [MeetingVariable.MeetingTypeShort]: meetingType.replace(/[^A-Z]/g, ''),
    [AdministrativeVariable.President]: president?.toString(),
    [AdministrativeVariable.VicePresident]: vicePresident?.toString(),
    [AdministrativeVariable.Secretary]: secretary?.toString(),
  };

  new Transaction()
    .step({
      forward: () =>
        (meetingMinutes = minutesTemplate.makeCopy(substituteVariablesInString(minutesTemplate.getName(), templateVariables), targetDir)),
      backward: () => meetingMinutes?.setTrashed(true),
    })
    .step({
      forward: () => meetingMinutes.setName(substituteVariablesInString(meetingMinutes.getName(), templateVariables)),
    })
    .step({
      forward: () => substituteVariablesInFile(meetingMinutes, templateVariables),
    })
    .run();

  return meetingMinutes;
};

export const createMeetingMinutes = () =>
  SafeWrapper.factory(createMeetingMinutes.name, { allowedEmails: institutionalEmails }).wrap((logger: SheetLogger) => {
    const president = getDirector(SaDepartment.Presidency);
    const vicePresident = getDirector(SaDepartment.VicePresidency);
    const secretary = getDirector(SaDepartment.Secretary);
    const attendees = getAttendees();
    const meetingType = getNamedValue(NamedRange.MeetingType);

    const [isPresidentPresent, isVicePresidentPresent, isSecretaryPresent] = areElementsInList(
      [president, vicePresident, secretary],
      attendees,
      (a, b) => a?.nUsp === b?.nUsp && a?.nUsp !== null,
    );

    // usually:
    // secretary is the clerk if present
    // otherwise, vice president if the clerk if president is present
    const clerk =
      (isSecretaryPresent ? secretary : isVicePresidentPresent && isPresidentPresent ? vicePresident : null)?.toString() ?? '???';

    logger.log(DialogTitle.InProgress, `Execução iniciada para reunião com ${clerk} na redação.`);

    const meetingStart = new Date();
    const minutesFile = createMinutesFile(meetingType, meetingStart, attendees, president, vicePresident, secretary, clerk);

    // clear attendee checkboxes
    GS.ss.getRangeByName(NamedRange.MeetingAttendees).clearContent();

    const webhooks = [
      new DiscordWebhook(getNamedValue(NamedRange.WebhookGeneral)),
      new DiscordWebhook(getNamedValue(NamedRange.WebhookBoardOfDirectors)),
    ];
    const body = `Ata criada com sucesso:\n${minutesFile.getUrl()}`;

    alert({ title: DialogTitle.Success, body });
    logger.log(DialogTitle.Success, body);

    webhooks.forEach(webhook =>
      webhook.post({
        embeds: buildDiscordEmbeds(meetingType, minutesFile, meetingStart, attendees, president, vicePresident, secretary, clerk),
      }),
    );
    webhooks.some(webhook => webhook.url.isUrl()) && logger?.log(DialogTitle.InProgress, 'Ata enviada no Discord.');
  });
