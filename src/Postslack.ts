import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { format } from "date-fns";

import { getConfig } from "./config";
import { EventInfo } from "./shift-changer-api";

type PartTimerProfile = {
  job: string;
  lastName: string;
  email: string;
  managerEmails: string[];
};

export const createModificationMessage = (
  modificationInfos: {
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  }[],
  partTimerProfile: PartTimerProfile
): string | undefined => {
  const messages = modificationInfos.map(({ previousEventInfo, newEventInfo }) => {
    return `${createMessageFromEventInfo(previousEventInfo)}\n↓\n${createMessageFromEventInfo(newEventInfo)}`;
  });
  if (messages.length == 0) return;
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が変更されました。`;
  return `${messageTitle}\n${messages.join("\n\n")}`;
};

export const createDeletionMessage = (
  deletionInfos: EventInfo[],
  partTimerProfile: PartTimerProfile
): string | undefined => {
  const messages = deletionInfos.map(createMessageFromEventInfo);
  if (messages.length == 0) return;
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が削除されました。`;
  return `${messageTitle}\n${messages.join("\n")}`;
};

export const getSlackClient = (slackToken: string): SlackClient => {
  return new SlackClient(slackToken);
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
export const createTitleFromEventInfo = (
  eventInfo: {
    restStartTime?: string;
    restEndTime?: string;
    workingStyle: string;
  },
  partTimerProfile: PartTimerProfile
): string => {
  const { job, lastName } = partTimerProfile;

  const restStartTime = eventInfo.restStartTime;
  const restEndTime = eventInfo.restEndTime;
  const workingStyle = eventInfo.workingStyle;

  if (restStartTime === undefined || restEndTime === undefined) {
    const title = `【${workingStyle}】${job}${lastName}さん`;
    return title;
  } else {
    const title = `【${workingStyle}】${job}${lastName}さん (休憩: ${restStartTime}~${restEndTime})`;
    return title;
  }
};
