import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { format } from "date-fns";

import { getConfig } from "./config";
import { EventInfo } from "./shift-changer-api";
type SheetType = "registration" | "modificationAndDeletion";
type PartTimerProfile = {
  job: string;
  lastName: string;
  email: string;
  managerEmails: string[];
};
export const getPartTimerProfile = (userEmail: string): PartTimerProfile => {
  const { JOB_SHEET_URL } = getConfig();
  const sheet = SpreadsheetApp.openByUrl(JOB_SHEET_URL).getSheetByName("シート1");
  if (!sheet) throw new Error("SHEET is not defined");
  const partTimerProfiles = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues()
    .map((row) => ({
      job: row[0] as string,
      // \u3000は全角空白
      lastName: row[1].split(/(\s|\u3000)+/)[0] as string,
      email: row[2] as string,
      managerEmails: row[3] === "" ? [] : (row[3] as string).replaceAll(/\s/g, "").split(","),
    }));

  const partTimerProfile = partTimerProfiles.find(({ email }) => {
    return email === userEmail;
  });
  if (partTimerProfile === undefined) throw new Error("no part timer information for the email");

  return partTimerProfile;
};
const slackIdToMention = (slackId: string) => `<@${slackId}>`;
export const postMessageToSlackChannel = (
  client: SlackClient,
  slackChannelToPost: string,
  messageToNotify: string,
  partTimerProfile: PartTimerProfile
) => {
  const { HR_MANAGER_SLACK_ID } = getConfig();
  const { managerEmails } = partTimerProfile;
  const managerSlackIds = getManagerSlackIds(managerEmails, client);
  const mentionMessageToManagers = [HR_MANAGER_SLACK_ID, ...managerSlackIds].map(slackIdToMention).join(" ");
  client.chat.postMessage({
    channel: slackChannelToPost,
    text: `${mentionMessageToManagers}\n${messageToNotify}`,
  });
};
const getManagerSlackIds = (managerEmails: string[], client: SlackClient): string[] => {
  const slackMembers = client.users.list().members ?? [];

  const managerSlackIds = managerEmails
    .map((email) => {
      const member = slackMembers.find((slackMember) => {
        return slackMember.profile?.email === email;
      });
      if (member === undefined) throw new Error("The manager email is not in the slack members");
      return member.id;
    })
    .filter((id): id is string => id !== undefined);

  return managerSlackIds;
};
export const getSheet = (sheetType: SheetType, spreadsheetUrl: string): GoogleAppsScript.Spreadsheet.Sheet => {
  const sheet = SpreadsheetApp.openByUrl(spreadsheetUrl)
    .getSheets()
    .find((sheet) =>
      sheet.getDeveloperMetadata().some((metaData) => metaData.getKey() === `part-timer-shift-manager-${sheetType}`)
    );

  if (!sheet) throw new Error("SHEET is not defined");

  return sheet;
};
export const createMessageFromEventInfo = (eventInfo: EventInfo) => {
  const date = format(new Date(eventInfo.date), "MM/dd");
  const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(eventInfo.title);
  if (restStartTime === undefined || restEndTime === undefined)
    return `【${workingStyle}】 ${date} ${eventInfo.startTime}~${eventInfo.endTime}`;
  else
    return `【${workingStyle}】 ${date} ${eventInfo.startTime}~${eventInfo.endTime} (休憩: ${restStartTime}~${restEndTime})`;
};
const getEventInfoFromTitle = (
  title: string
): { workingStyle?: string; restStartTime?: string; restEndTime?: string } => {
  const workingStyleRegex = /【(.*?)】/;
  const matchResult = title.match(workingStyleRegex)?.[1];
  const workingStyle = matchResult ?? "未設定";

  const restTimeRegex = /\d{2}:\d{2}~\d{2}:\d{2}/;
  const restTimeResult = title.match(restTimeRegex)?.[0];
  const [restStartTime, restEndTime] = restTimeResult ? restTimeResult.split("~") : [];
  return { workingStyle, restStartTime, restEndTime };
};
export const getSlackClient = (slackToken: string): SlackClient => {
  return new SlackClient(slackToken);
};
export const createTitleFromEventInfo = (
  eventInfo: {
    restStartTime?: Date;
    restEndTime?: Date;
    workingStyle: string;
  },
  partTimerProfile: PartTimerProfile
): string => {
  const { job, lastName } = partTimerProfile;

  const restStartTime = format(eventInfo.restStartTime as Date, "HH:mm");
  const restEndTime = format(eventInfo.restEndTime as Date, "HH:mm");
  const workingStyle = eventInfo.workingStyle;

  if (restStartTime === undefined || restEndTime === undefined) {
    const title = `【${workingStyle}】${job}${lastName}さん`;
    return title;
  } else {
    const title = `【${workingStyle}】${job}${lastName}さん (休憩: ${restStartTime}~${restEndTime})`;
    return title;
  }
};
