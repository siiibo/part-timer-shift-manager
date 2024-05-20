import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { format } from "date-fns";

import { getConfig } from "./config";
import { PartTimerProfile } from "./JobSheet";
import { getPartTimerProfile } from "./JobSheet";
import {
  getModificationOrDeletion,
  insertModificationAndDeletionSheet,
  setValuesModificationAndDeletionSheet,
} from "./ModificationAndDeletionSheet";
import {
  DeletionRecurringEventRow,
  getRecurringEventModificationOrDeletionOrRegistration,
  insertRecurringEventSheet,
  ModificationRecurringEventRow,
  RegistrationRecurringEventRow,
} from "./RecurringEventSheet";
import { getRegistrationRows, insertRegistrationSheet, setValuesRegistrationSheet } from "./RegistrationSheet";
import { EventInfo, shiftChanger } from "./shift-changer-api";

type RecurringType = RegistrationRecurringEventRow[] | ModificationRecurringEventRow[] | DeletionRecurringEventRow[];
type SheetType = "registration" | "modificationAndDeletion" | "recurringEvent";
type OperationType = "registration" | "modificationAndDeletion" | "showEvents" | "recurringEvent";
export const doGet = () => {
  return ContentService.createTextOutput("ok");
};
export const doPost = (e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput => {
  if (e.parameter.apiId === "shift-changer") {
    const response = shiftChanger(e) ?? "";
    return ContentService.createTextOutput(response).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput("undefined");
};

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
        .createMenu("登録")
        .addItem("シートの追加", insertRegistrationSheet.name)
        .addSeparator()
        .addItem("提出", callRegistration.name),
    )
    .addSubMenu(
      ui
        .createMenu("変更・削除")
        .addItem("シートの追加", insertModificationAndDeletionSheet.name)
        .addSeparator()
        .addItem("予定を表示", callShowEvents.name)
        .addItem("提出", callModificationAndDeletion.name),
    )
    .addSubMenu(
      ui
        .createMenu("固定シフト変更")
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

  const sheetType: SheetType = "registration";
  const sheet = getSheet(sheetType, spreadsheetUrl);
  const operationType: OperationType = "registration";
  const comment = sheet.getRange("A2").getValue();
  const registrationRows = getRegistrationRows(sheet);
  const registrationInfos = registrationRows.map((registrationRow) => {
    const title = createTitleFromEventInfo(
      {
        ...(registrationRow.restStartTime && { restStartTime: registrationRow.restStartTime }),
        ...(registrationRow.restEndTime && { restEndTime: registrationRow.restEndTime }),
        workingStyle: registrationRow.workingStyle,
      },
      partTimerProfile,
    );
    return {
      title: title,
      date: registrationRow.startTime,
      startTime: registrationRow.startTime,
      endTime: registrationRow.endTime,
    };
  });
  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    registrationInfos: JSON.stringify(registrationInfos),
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
  registrationInfos: EventInfo[],
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
  const sheetType: SheetType = "modificationAndDeletion";
  const sheet = getSheet(sheetType, spreadsheetUrl);
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
  const eventInfos = EventInfo.array().parse(JSON.parse(response.getContentText()));

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
  const sheetType: SheetType = "modificationAndDeletion";
  const sheet = getSheet(sheetType, spreadsheetUrl);
  const comment = sheet.getRange("A2").getValue();
  const operationType: OperationType = "modificationAndDeletion";
  const { modificationRows, deletionRows } = getModificationOrDeletion(sheet);
  const modificationInfos = modificationRows.map((modificationRow) => {
    const newTitle = createTitleFromEventInfo(
      {
        ...(modificationRow.newRestStartTime && { restStartTime: modificationRow.newRestStartTime }),
        ...(modificationRow.newRestEndTime && { restEndTime: modificationRow.newRestEndTime }),
        workingStyle: modificationRow.newWorkingStyle,
      },
      partTimerProfile,
    );
    return {
      previousEventInfo: {
        title: modificationRow.title,
        date: modificationRow.startTime,
        startTime: modificationRow.startTime,
        endTime: modificationRow.endTime,
      },
      newEventInfo: {
        title: newTitle,
        date: modificationRow.newStartTime,
        startTime: modificationRow.newStartTime,
        endTime: modificationRow.newEndTime,
      },
    };
  });
  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    modificationInfos: JSON.stringify(modificationInfos),
    deletionInfos: JSON.stringify(deletionRows),
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
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  }[],
  partTimerProfile: PartTimerProfile,
): string | undefined => {
  const messages = modificationInfos.map(({ previousEventInfo, newEventInfo }) => {
    return `${createMessageFromEventInfo(previousEventInfo)}\n↓\n${createMessageFromEventInfo(newEventInfo)}`;
  });
  if (messages.length == 0) return;
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が変更されました。`;
  return `${messageTitle}\n${messages.join("\n\n")}`;
};
const createDeletionMessage = (deletionInfos: EventInfo[], partTimerProfile: PartTimerProfile): string | undefined => {
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
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const { SLACK_ACCESS_TOKEN } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  const partTimerProfile = getPartTimerProfile(userEmail);
  const sheetType: SheetType = "recurringEvent";
  const sheet = getSheet(sheetType, spreadsheetUrl);
  const comment = sheet.getRange("A2").getValue();
  const { registrationRows, modificationRows, deletionRows } =
    getRecurringEventModificationOrDeletionOrRegistration(sheet);
  const after = new Date(sheet.getRange("A5").getValue());

  const registrationInfos = registrationRows.map((registrationRow) => {
    const title = createTitleFromEventInfo(
      {
        ...(registrationRow.restStartTime && { restStartTime: registrationRow.restStartTime }),
        ...(registrationRow.restEndTime && { restEndTime: registrationRow.restEndTime }),
        workingStyle: registrationRow.workingStyle,
      },
      partTimerProfile,
    );
    return {
      title: title,
      dayOfWeek: registrationRow.dayOfWeek,
      startTime: registrationRow.startTime,
      endTime: registrationRow.endTime,
    };
  });
  const modificationInfos = modificationRows.map((modificationRow) => {
    const title = createTitleFromEventInfo(
      {
        ...(modificationRow.restStartTime && { restStartTime: modificationRow.restStartTime }),
        ...(modificationRow.restEndTime && { restEndTime: modificationRow.restEndTime }),
        workingStyle: modificationRow.workingStyle,
      },
      partTimerProfile,
    );
    return {
      title: title,
      dayOfWeek: modificationRow.dayOfWeek,
      startTime: modificationRow.startTime,
      endTime: modificationRow.endTime,
    };
  });

  const deleteDayOfWeeks = deletionRows.map((deletionRow) => {
    return deletionRow.dayOfWeek;
  });
  if (modificationInfos.length == 0 && deletionRows.length == 0 && registrationInfos.length == 0) {
    throw new Error("追加・変更・削除する予定がありません。");
  }

  const { API_URL, SLACK_CHANNEL_TO_POST } = getConfig();
  if (registrationInfos.length > 0) {
    const payload = {
      apiId: "shift-changer",
      operationType: "registerRecurringEvent",
      userEmail: userEmail,
      recurringEventRegistration: JSON.stringify({ after, events: registrationInfos }),
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(response.getContentText());
    }
  }
  if (modificationInfos.length > 0) {
    const payload = {
      apiId: "shift-changer",
      operationType: "modifyRecurringEvent",
      userEmail: userEmail,
      recurringEventModification: JSON.stringify({ after, events: modificationInfos }),
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(response.getContentText());
    }
  }
  if (deleteDayOfWeeks.length > 0) {
    const payload = {
      apiId: "shift-changer",
      operationType: "deleteRecurringEvent",
      userEmail: userEmail,
      recurringEventDeletion: JSON.stringify({ after, dayOfWeeks: deleteDayOfWeeks }),
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(response.getContentText());
    }
  }

  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の繰り返し予定が変更されました`;
  const RecurringEventMessageToNotify = [
    `${messageTitle}`,
    createMessageForRecurringEvent(registrationRows),
    createMessageForRecurringEvent(modificationRows),
    createMessageForRecurringEvent(deletionRows),
    comment ? `コメント: ${comment}` : undefined,
  ]
    .filter(Boolean)
    .join("\n---\n");

  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, RecurringEventMessageToNotify, partTimerProfile);
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
const createMessageFromEventInfo = (eventInfo: EventInfo) => {
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

const createMessageForRecurringEvent = (recurringEventInfos: RecurringType): string => {
  const messageTitle = {
    modification: "以下の繰り返し予定が変更されました",
    registration: "以下の繰り返し予定が追加されました",
    deletion: "以下の繰り返し予定が削除されました",
  };
  if (recurringEventInfos.length === 0) return "";

  const messages = recurringEventInfos.map((recurringEventInfo) => {
    if (recurringEventInfo.type === "registration") {
      const { dayOfWeek, startTime, endTime, restStartTime, restEndTime, workingStyle } = recurringEventInfo;

      if (restStartTime && restEndTime) {
        return `${dayOfWeek} : ${workingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")} (休憩: ${format(restStartTime, "HH:mm")}~${format(restEndTime, "HH:mm")})`;
      } else {
        return `${dayOfWeek} : ${workingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`;
      }
    } else if (recurringEventInfo.type === "modification") {
      const { dayOfWeek, startTime, endTime, restStartTime, restEndTime, workingStyle } = recurringEventInfo;

      if (restEndTime && restStartTime) {
        return `${dayOfWeek} : ${workingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")} (休憩: ${format(restStartTime, "HH:mm")}~${format(restEndTime, "HH:mm")})`;
      } else {
        return `${dayOfWeek} : ${workingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`;
      }
    } else if (recurringEventInfo && recurringEventInfo.type === "deletion") {
      const { dayOfWeek } = recurringEventInfo;

      return `${dayOfWeek}`;
    }
  });
  return `${messageTitle[recurringEventInfos[0].type]}\n${messages.join("\n")}`;
};
