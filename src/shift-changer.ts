import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { format } from "date-fns";

import { DayOfWeek } from "./common.schema";
import { getConfig } from "./config";
import { PartTimerProfile } from "./JobSheet";
import { getPartTimerProfile } from "./JobSheet";
import {
  getModificationOrDeletion,
  insertModificationAndDeletionSheet,
  setValuesModificationAndDeletionSheet,
} from "./ModificationAndDeletionSheet";
import {
  getRecurringEventSheetValues,
  insertRecurringEventSheet,
  setValuesRecurringEventSheet,
} from "./RecurringEventSheet";
import { getRegistrationRows, insertRegistrationSheet, setValuesRegistrationSheet } from "./RegistrationSheet";
import { Event, OperationType, RecurringEventResponse } from "./shift-changer-api";

type SheetType = "registration" | "modificationAndDeletion" | "recurringEvent";

export const initShiftChanger = () => {
  const { DEV_SPREADSHEET_URL } = getConfig();
  ScriptApp.newTrigger(onOpenForDev.name)
    .forSpreadsheet(SpreadsheetApp.openByUrl(DEV_SPREADSHEET_URL))
    .onOpen()
    .create();
};

export const onOpen = () => {
  const ui = SpreadsheetApp.getUi();
  createMenu(ui, ui.createAddonMenu());
};

export const onOpenForDev = () => {
  const ui = SpreadsheetApp.getUi();
  createMenu(ui, ui.createMenu("[dev] シフト変更ツール"));
};

const createMenu = (ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) => {
  return menu
    .addSubMenu(
      ui
        .createMenu("シフト登録")
        .addItem("シートの追加", insertRegistrationSheet.name)
        .addSeparator()
        .addItem("提出", callRegistration.name),
    )
    .addSubMenu(
      ui
        .createMenu("シフト変更・削除")
        .addItem("シートの追加", insertModificationAndDeletionSheet.name)
        .addSeparator()
        .addItem("予定を表示", callShowEvents.name)
        .addItem("提出", callModificationAndDeletion.name),
    )
    .addSubMenu(
      ui
        .createMenu("固定シフト登録・変更・消去")
        .addItem("シートの追加", insertRecurringEventSheet.name)
        .addSeparator()
        .addItem("提出", callRecurringEvent.name),
    )
    .addToUi();
};

export const callRegistration = () => {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("すでに処理を実行中です。そのままお待ちください");
  }
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const { SLACK_ACCESS_TOKEN } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  const partTimerProfile = getPartTimerProfile(userEmail);

  const sheet = getSheet("registration", spreadsheetUrl);
  const operationType: OperationType = "registerEvent";
  const comment = sheet.getRange("A2").getValue();
  const registrationRows = getRegistrationRows(sheet);
  const registrationInfos = registrationRows.map(({ startTime, endTime, restStartTime, restEndTime, workingStyle }) => {
    const title = createTitleFromEventInfo(
      {
        ...(restStartTime && { restStartTime: restStartTime }),
        ...(restEndTime && { restEndTime: restEndTime }),
        workingStyle,
      },
      partTimerProfile,
    );
    return { title, date: startTime, startTime, endTime };
  });
  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    registrationEvents: JSON.stringify(registrationInfos),
  };
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true,
  };
  const { API_URL, SLACK_CHANNEL_TO_POST } = getConfig();
  const response = UrlFetchApp.fetch(API_URL, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(response.getContentText());
  }
  const messageToNotify = createRegistrationMessage(registrationInfos, comment, partTimerProfile);
  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, messageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesRegistrationSheet(sheet);
};

const createRegistrationMessage = (
  registrationInfos: Event[],
  comment: string,
  partTimerProfile: PartTimerProfile,
): string => {
  const messages = registrationInfos.map(createMessageFromEventInfo);
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が追加されました。`;
  return comment
    ? `${messageTitle}\n${messages.join("\n")}\n\nコメント: ${comment}`
    : `${messageTitle}\n${messages.join("\n")}`;
};

export const callShowEvents = () => {
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheet = getSheet("modificationAndDeletion", spreadsheetUrl);
  const operationType: OperationType = "showEvents";
  const startDate: Date = sheet.getRange("A5").getValue();
  if (!startDate) throw new Error("日付を指定してください。");

  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    startDate: startDate,
  };
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true,
  };
  const { API_URL } = getConfig();
  const response = UrlFetchApp.fetch(API_URL, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(response.getContentText());
  }
  const eventInfos = Event.array().parse(JSON.parse(response.getContentText()));

  if (eventInfos.length === 0) throw new Error("no events");

  const moldedEventInfos = eventInfos.map(({ title, date, startTime, endTime }) => {
    const dateStr = format(date, "yyyy/MM/dd");
    const startTimeStr = format(startTime, "HH:mm");
    const endTimeStr = format(endTime, "HH:mm");
    return [title, dateStr, startTimeStr, endTimeStr];
  });

  if (sheet.getLastRow() > 8) {
    sheet.getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn()).clearContent();
  }

  sheet.getRange(9, 1, moldedEventInfos.length, moldedEventInfos[0].length).setValues(moldedEventInfos);
};

export const callModificationAndDeletion = () => {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("すでに処理を実行中です。そのままお待ちください");
  }
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const { SLACK_ACCESS_TOKEN } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  const partTimerProfile = getPartTimerProfile(userEmail);
  const sheet = getSheet("modificationAndDeletion", spreadsheetUrl);
  const comment = sheet.getRange("A2").getValue();
  const operationType: OperationType = "modifyAndDeleteEvent";
  const { modificationRows, deletionRows } = getModificationOrDeletion(sheet);
  const modificationInfos = modificationRows.map(
    ({ title, startTime, endTime, newStartTime, newEndTime, newRestStartTime, newRestEndTime, newWorkingStyle }) => {
      const newTitle = createTitleFromEventInfo(
        {
          ...(newRestStartTime && { restStartTime: newRestStartTime }),
          ...(newRestEndTime && { restEndTime: newRestEndTime }),
          workingStyle: newWorkingStyle,
        },
        partTimerProfile,
      );
      return {
        previousEvent: { title, date: startTime, startTime, endTime },
        newEvent: {
          title: newTitle,
          date: newStartTime,
          startTime: newStartTime,
          endTime: newEndTime,
        },
      };
    },
  );
  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    modificationEvents: JSON.stringify(modificationInfos),
    deletionEvents: JSON.stringify(deletionRows),
  };
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true,
  };
  const { API_URL, SLACK_CHANNEL_TO_POST } = getConfig();
  const response = UrlFetchApp.fetch(API_URL, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(response.getContentText());
  }
  if (modificationInfos.length == 0 && deletionRows.length == 0) {
    throw new Error("変更・削除する予定がありません。");
  }
  const modificationAndDeletionMessageToNotify = [
    createModificationMessage(modificationInfos, partTimerProfile),
    createDeletionMessage(deletionRows, partTimerProfile),
    comment ? `コメント: ${comment}` : undefined,
  ]
    .filter(Boolean)
    .join("\n---\n");

  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, modificationAndDeletionMessageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesModificationAndDeletionSheet(sheet);
};

const createModificationMessage = (
  modificationInfos: {
    previousEvent: Event;
    newEvent: Event;
  }[],
  partTimerProfile: PartTimerProfile,
): string | undefined => {
  const messages = modificationInfos.map(({ previousEvent, newEvent }) => {
    return `${createMessageFromEventInfo(previousEvent)}\n↓\n${createMessageFromEventInfo(newEvent)}`;
  });
  if (messages.length == 0) return;
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が変更されました。`;
  return `${messageTitle}\n${messages.join("\n\n")}`;
};
const createDeletionMessage = (deletionInfos: Event[], partTimerProfile: PartTimerProfile): string | undefined => {
  const messages = deletionInfos.map(createMessageFromEventInfo);
  if (messages.length == 0) return;
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が削除されました。`;
  return `${messageTitle}\n${messages.join("\n")}`;
};

export const callRecurringEvent = () => {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("すでに処理を実行中です。そのままお待ちください");
  }
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheet = getSheet("recurringEvent", spreadsheetUrl);
  const { after, comment, registrationRows, modificationRows, deletionRows } = getRecurringEventSheetValues(sheet);
  const userEmail = Session.getActiveUser().getEmail();
  const partTimerProfile = getPartTimerProfile(userEmail);

  if (registrationRows.length == 0 && modificationRows.length == 0 && deletionRows.length == 0) {
    throw new Error("追加・変更・削除する予定がありません。");
  }

  const registrationInfos = registrationRows.map(
    ({ startTime, endTime, restStartTime, restEndTime, dayOfWeek, workingStyle }) => {
      const title = createTitleFromEventInfo(
        {
          ...(restStartTime && { restStartTime }),
          ...(restEndTime && { restEndTime }),
          workingStyle,
        },
        partTimerProfile,
      );
      return { title, dayOfWeek, startTime, endTime };
    },
  );
  const modificationInfos = modificationRows.map(
    ({ startTime, endTime, restStartTime, restEndTime, dayOfWeek, workingStyle }) => {
      const title = createTitleFromEventInfo(
        {
          ...(restStartTime && { restStartTime }),
          ...(restEndTime && { restEndTime }),
          workingStyle,
        },
        partTimerProfile,
      );
      return { title, dayOfWeek, startTime, endTime };
    },
  );

  const deleteDayOfWeeks = deletionRows.map((deletionRow) => {
    return deletionRow.dayOfWeek;
  });
  const deleteInfos = deleteDayOfWeeks.map((deletionRow) => {
    return { dayOfWeek: deletionRow };
  });

  const payloadBase = { apiId: "shift-changer", userEmail };
  const { API_URL } = getConfig();
  if (registrationInfos.length > 0) {
    const payload = {
      ...payloadBase,
      operationType: "registerRecurringEvent",
      registrationRecurringEvents: JSON.stringify({ after, events: registrationInfos }),
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    UrlFetchApp.fetch(API_URL, options);
  }
  if (modificationInfos.length > 0) {
    const payload = {
      ...payloadBase,
      operationType: "modifyRecurringEvent",
      modificationRecurringEvents: JSON.stringify({ after, events: modificationInfos }),
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseContent = RecurringEventResponse.parse(JSON.parse(response.getContentText()));
    if (responseContent.error) {
      //NOTE: APIのレスポンスがある場合はエラーを出力する
      throw new Error(responseContent.error);
    }
  }
  if (deleteDayOfWeeks.length > 0) {
    const payload = {
      ...payloadBase,
      operationType: "deleteRecurringEvent",
      deletionRecurringEvents: JSON.stringify({ after, dayOfWeeks: deleteDayOfWeeks }),
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseContent = RecurringEventResponse.parse(JSON.parse(response.getContentText()));
    if (responseContent.error) {
      //NOTE: APIのレスポンスがある場合はエラーを出力する
      throw new Error(responseContent.error);
    }
  }

  const recurringEventMessageToNotify = createMessageForRecurringEvent(
    after,
    partTimerProfile,
    registrationInfos,
    modificationInfos,
    deleteInfos,
    comment,
  );

  const { SLACK_ACCESS_TOKEN, SLACK_CHANNEL_TO_POST } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, recurringEventMessageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesRecurringEventSheet(sheet);
};

const createMessageForRecurringEvent = (
  after: Date,
  { job, lastName }: PartTimerProfile,
  registrationInfos: { title: string; dayOfWeek: DayOfWeek; startTime: Date; endTime: Date }[],
  modificationInfos: { title: string; dayOfWeek: DayOfWeek; startTime: Date; endTime: Date }[],
  deletionInfos: { dayOfWeek: DayOfWeek }[],
  comment: string | undefined,
): string => {
  const registrationMessages = registrationInfos.map(
    ({ title, dayOfWeek, startTime, endTime }) =>
      `${dayOfWeek} : ${title} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`,
  );
  const modificationMessages = modificationInfos.map(
    ({ title, dayOfWeek, startTime, endTime }) =>
      `${dayOfWeek} : ${title} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`,
  );
  const deletionMessages = deletionInfos.map(({ dayOfWeek }) => `${dayOfWeek}`).join(", ");

  const message = [
    `${format(after, "yyyy/MM/dd")}以降の繰り返し予定を変更しました`,
    registrationMessages.length > 0
      ? `${job}${lastName}さんの以下の繰り返し予定が追加されました\n${registrationMessages.join("\n")}`
      : "",
    modificationMessages.length > 0
      ? `${job}${lastName}さんの以下の繰り返し予定が変更されました\n${modificationMessages.join("\n")}`
      : "",
    deletionMessages.length > 0 ? `${job}${lastName}さんの以下の繰り返し予定が削除されました\n${deletionMessages}` : "",
  ]
    .filter(Boolean)
    .join("\n---\n");

  return comment ? `${message}\n---\nコメント: ${comment}` : message;
};

const getSlackClient = (slackToken: string): SlackClient => {
  return new SlackClient(slackToken);
};

const getSheet = (sheetType: SheetType, spreadsheetUrl: string): GoogleAppsScript.Spreadsheet.Sheet => {
  const sheet = SpreadsheetApp.openByUrl(spreadsheetUrl)
    .getSheets()
    .find((sheet) =>
      sheet.getDeveloperMetadata().some((metaData) => metaData.getKey() === `part-timer-shift-manager-${sheetType}`),
    );

  if (!sheet) throw new Error("SHEET is not defined");

  return sheet;
};

const createTitleFromEventInfo = (
  eventInfo: {
    restStartTime?: Date;
    restEndTime?: Date;
    workingStyle: string;
  },
  partTimerProfile: PartTimerProfile,
): string => {
  const { job, lastName } = partTimerProfile;

  const restStartTime = eventInfo.restStartTime ? format(eventInfo.restStartTime, "HH:mm") : undefined;
  const restEndTime = eventInfo.restEndTime ? format(eventInfo.restEndTime, "HH:mm") : undefined;
  const workingStyle = eventInfo.workingStyle;

  if (restStartTime === undefined || restEndTime === undefined) {
    const title = `【${workingStyle}】${job}${lastName}さん`;
    return title;
  } else {
    const title = `【${workingStyle}】${job}${lastName}さん (休憩: ${restStartTime}~${restEndTime})`;
    return title;
  }
};

const createMessageFromEventInfo = (eventInfo: Event) => {
  const date = format(eventInfo.date, "MM/dd");
  const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(eventInfo.title);
  const startTime = format(eventInfo.startTime, "HH:mm");
  const endTime = format(eventInfo.endTime, "HH:mm");
  if (restStartTime === undefined || restEndTime === undefined)
    return `【${workingStyle}】 ${date} ${startTime}~${endTime}`;
  else return `【${workingStyle}】 ${date} ${startTime}~${endTime} (休憩: ${restStartTime}~${restEndTime})`;
};
const getEventInfoFromTitle = (
  title: string,
): { workingStyle?: string; restStartTime?: string; restEndTime?: string } => {
  const workingStyleRegex = /【(.*?)】/;
  const matchResult = title.match(workingStyleRegex)?.[1];
  const workingStyle = matchResult ?? "未設定";

  const restTimeRegex = /\d{2}:\d{2}~\d{2}:\d{2}/;
  const restTimeResult = title.match(restTimeRegex)?.[0];
  const [restStartTime, restEndTime] = restTimeResult ? restTimeResult.split("~") : [];
  return { workingStyle, restStartTime, restEndTime };
};

const slackIdToMention = (slackId: string) => `<@${slackId}>`;
const postMessageToSlackChannel = (
  client: SlackClient,
  slackChannelToPost: string,
  messageToNotify: string,
  partTimerProfile: PartTimerProfile,
) => {
  const { HR_MANAGER_SLACK_ID } = getConfig();
  const { managerEmails } = partTimerProfile;
  const managerSlackIds = getManagerSlackIds(managerEmails, client);
  const mentionMessageToManagers = HR_MANAGER_SLACK_ID ? [HR_MANAGER_SLACK_ID, ...managerSlackIds] : managerSlackIds;
  const mentionMessage = mentionMessageToManagers.map(slackIdToMention).join(" ");
  client.chat.postMessage({
    channel: slackChannelToPost,
    text: `${mentionMessage}\n${messageToNotify}`,
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
