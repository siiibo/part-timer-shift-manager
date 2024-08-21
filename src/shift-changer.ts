import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { format } from "date-fns";
import { z } from "zod";

import { deleteHolidayShift } from "./autoDeleteHolidayEvent";
import { DayOfWeek } from "./common.schema";
import { getConfig } from "./config";
import { PartTimerProfile } from "./JobSheet";
import { getPartTimerProfile } from "./JobSheet";
import {
  getModificationOrDeletionSheetValues,
  insertModificationAndDeletionSheet,
  setValuesModificationAndDeletionSheet,
} from "./ModificationAndDeletionSheet";
import {
  getRecurringEventSheetValues,
  insertRecurringEventSheet,
  setValuesRecurringEventSheet,
} from "./RecurringEventSheet";
import { getRegistrationSheetValues, insertRegistrationSheet, setValuesRegistrationSheet } from "./RegistrationSheet";
import {
  APIResponse,
  DeleteEventRequest,
  DeleteRecurringEventRequest,
  Event,
  ModifyEventRequest,
  ModifyRecurringEventRequest,
  RegisterEventRequest,
  RegisterRecurringEventRequest,
  ShowEventRequest,
} from "./shift-changer-api";

const CreateMessageSchema = z.union([
  z.object({
    type: z.literal("registerEvent"),
    eventInfos: z.array(Event),
  }),
  z.object({
    type: z.literal("modifyEvent"),
    eventInfos: z
      .object({
        previousEvent: Event,
        newEvent: Event,
      })
      .array(),
  }),
  z.object({
    type: z.literal("deleteEvent"),
    eventInfos: z.array(Event),
  }),
]);
type CreateMessageSchema = z.infer<typeof CreateMessageSchema>;

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
        .createMenu("単発シフト登録")
        .addItem("シートの追加", insertRegistrationSheet.name)
        .addSeparator()
        .addItem("提出", callRegistration.name),
    )
    .addSubMenu(
      ui
        .createMenu("単発シフト変更・削除")
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
  const { comment, registrationRows } = getRegistrationSheetValues(sheet);
  console.info(`registration: ${JSON.stringify(registrationRows)}`); //NOTE: シート内容を確認するためのログ
  const registrationInfos = registrationRows.map(({ startTime, endTime, restStartTime, restEndTime, workingStyle }) => {
    const title = createTitleFromEventInfo(
      {
        ...(restStartTime && { restStartTime: restStartTime }),
        ...(restEndTime && { restEndTime: restEndTime }),
        workingStyle,
      },
      partTimerProfile,
    );
    return { title, startTime, endTime };
  });

  const basePayload = { apiId: "shift-changer", userEmail: userEmail } as const;
  const payload = JSON.stringify({
    ...basePayload,
    operationType: "registerEvent",
    events: registrationInfos,
  } satisfies RegisterEventRequest);
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true,
  };
  const { API_URL, SLACK_CHANNEL_TO_POST } = getConfig();
  UrlFetchApp.fetch(API_URL, options);
  const { job, lastName } = partTimerProfile;
  const messageToNotify = [
    `${job}${lastName}さんが以下の単発シフトを変更しました`,
    createMessage({ type: "registerEvent", eventInfos: registrationInfos }),
    "---",
    `コメント: ${comment}`,
  ].join("\n");
  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, messageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesRegistrationSheet(sheet);
};

export const callShowEvents = () => {
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheet = getSheet("modificationAndDeletion", spreadsheetUrl);
  const startDate: Date = sheet.getRange("A5").getValue();
  if (!startDate) throw new Error("日付を指定してください。");

  const basePayload = { apiId: "shift-changer", userEmail: userEmail } as const;
  const payload = JSON.stringify({
    ...basePayload,
    operationType: "showEvents",
    startDate: startDate,
  } satisfies ShowEventRequest);
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true,
  };
  const { API_URL } = getConfig();
  const response = UrlFetchApp.fetch(API_URL, options);
  const responseContent = APIResponse.parse(JSON.parse(response.getContentText()));
  if ("error" in responseContent) {
    throw new Error(responseContent.error);
  }
  if (responseContent.events.length === 0) {
    throw new Error("no events");
  }

  const eventInfos = responseContent.events.map(({ title, startTime, endTime }) => {
    const dateStr = format(startTime, "yyyy/MM/dd");
    const startTimeStr = format(startTime, "HH:mm");
    const endTimeStr = format(endTime, "HH:mm");
    return [title, dateStr, startTimeStr, endTimeStr];
  });

  if (sheet.getLastRow() > 8) {
    sheet.getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn()).clearContent();
  }

  sheet.getRange(9, 1, eventInfos.length, eventInfos[0].length).setValues(eventInfos);
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

  const { comment, modificationRows, deletionRows } = getModificationOrDeletionSheetValues(sheet);
  console.info(`modification: ${JSON.stringify(modificationRows)},deletion: ${JSON.stringify(deletionRows)}`); //NOTE: シート内容を確認するためのログ
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
        previousEvent: { title, startTime, endTime },
        newEvent: {
          title: newTitle,
          startTime: newStartTime,
          endTime: newEndTime,
        },
      };
    },
  );

  const deleteInfos = deletionRows.map(({ title, startTime, endTime }) => {
    return { title, startTime, endTime };
  });

  const { API_URL } = getConfig();
  const basePayload = { apiId: "shift-changer", userEmail: userEmail } as const;
  if (modificationInfos.length > 0) {
    const payload = JSON.stringify({
      ...basePayload,
      operationType: "modifyEvent",
      events: modificationInfos,
    } satisfies ModifyEventRequest);
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    UrlFetchApp.fetch(API_URL, options);
  }
  if (deleteInfos.length > 0) {
    const payload = JSON.stringify({
      ...basePayload,
      operationType: "deleteEvent",
      events: deleteInfos,
    } satisfies DeleteEventRequest);
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    UrlFetchApp.fetch(API_URL, options);
  }

  const { SLACK_CHANNEL_TO_POST } = getConfig();
  const { job, lastName } = partTimerProfile;
  const modificationAndDeletionMessageToNotify = [
    `${job}${lastName}さんが以下の単発シフトを変更しました`,
    createMessage({ type: "modifyEvent", eventInfos: modificationInfos }),
    createMessage({ type: "deleteEvent", eventInfos: deleteInfos }),
    "---",
    `コメント: ${comment}`,
  ].join("\n");

  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, modificationAndDeletionMessageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesModificationAndDeletionSheet(sheet);
};

const createMessage = (messageInfos: CreateMessageSchema) => {
  if (messageInfos.type === "registerEvent") {
    const messages = messageInfos.eventInfos.map(createMessageFromEventInfo);
    const messageTitle = "[追加]";
    return `${messageTitle}\n${messages.join("\n")}`;
  }
  if (messageInfos.type === "modifyEvent") {
    const messages = messageInfos.eventInfos.map(({ previousEvent, newEvent }) => {
      return `${createMessageFromEventInfo(previousEvent)} → ${createMessageFromEventInfo(newEvent)}`;
    });
    const messageTitle = "[変更]";
    return `${messageTitle}\n${messages.join("\n")}`;
  }
  if (messageInfos.type === "deleteEvent") {
    const messages = messageInfos.eventInfos.map(createMessageFromEventInfo);
    const messageTitle = "[消去]";
    return `${messageTitle}\n${messages.join("\n")}`;
  }
};

export const callRecurringEvent = () => {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("すでに処理を実行中です。そのままお待ちください");
  }
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheet = getSheet("recurringEvent", spreadsheetUrl);
  const { after, comment, registrationRows, modificationRows, deletionRows } = getRecurringEventSheetValues(sheet);
  console.info(
    `recurringRegistration: ${JSON.stringify(registrationRows)},recurringModification: ${JSON.stringify(modificationRows)},recurringDeletion: ${JSON.stringify(deletionRows)}`,
  ); //NOTE: シート内容を確認するためのログ
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

  const basePayload = { apiId: "shift-changer", userEmail: userEmail } as const;
  const { API_URL } = getConfig();
  if (registrationInfos.length > 0) {
    const payload = JSON.stringify({
      ...basePayload,
      operationType: "registerRecurringEvent",
      recurringInfo: { after, events: registrationInfos },
    } satisfies RegisterRecurringEventRequest);
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    UrlFetchApp.fetch(API_URL, options);
  }
  let modifyEventStrings = ""; //NOTE: 繰り返し予定の変更APIの情報を利用するため、letで宣言
  if (modificationInfos.length > 0) {
    const payload = JSON.stringify({
      ...basePayload,
      operationType: "modifyRecurringEvent",
      recurringInfo: { after, events: modificationInfos },
    } satisfies ModifyRecurringEventRequest);
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseContent = APIResponse.parse(JSON.parse(response.getContentText()));
    if ("error" in responseContent) {
      //NOTE: APIのレスポンスがある場合はエラーを出力する
      throw new Error(responseContent.error);
    }
    modifyEventStrings = createMessageForModifyRecurringEvent(responseContent?.events, modificationInfos);
  }
  let deleteEventStrings = ""; //NOTE: 繰り返し予定の削除APIの情報を利用するため、letで宣言
  if (deleteDayOfWeeks.length > 0) {
    const payload = JSON.stringify({
      ...basePayload,
      operationType: "deleteRecurringEvent",
      recurringInfo: { after, dayOfWeeks: deleteDayOfWeeks },
    } satisfies DeleteRecurringEventRequest);
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseContent = APIResponse.parse(JSON.parse(response.getContentText()));
    if ("error" in responseContent) {
      //NOTE: APIのレスポンスがある場合はエラーを出力する
      throw new Error(responseContent.error);
    }
    deleteEventStrings = createMessageForDeleteRecurringEvent(responseContent?.events, deleteDayOfWeeks);
  }
  const recurringEventMessageToNotify = createMessageForRecurringEvent(
    partTimerProfile,
    after,
    createMessageForRegisterRecurringEvent(registrationInfos),
    modifyEventStrings,
    deleteEventStrings,
    comment,
  );

  const { SLACK_ACCESS_TOKEN, SLACK_CHANNEL_TO_POST } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, recurringEventMessageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesRecurringEventSheet(sheet);
  //NOTE: 繰り返し予定の入力制限ができないため、deleteHolidayShiftを実行して祝日の予定を削除する
  // ref: https://github.com/siiibo/part-timer-shift-manager/pull/53#discussion_r1665084529
  deleteHolidayShift();
};

const createMessageForRegisterRecurringEvent = (
  registrationInfos: { title: string; dayOfWeek: DayOfWeek; startTime: Date; endTime: Date }[],
): string => {
  if (registrationInfos.length === 0) return "";
  const messages = registrationInfos.map(({ title, dayOfWeek, startTime, endTime }) => {
    const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(title);
    const emojiWorkingStyle = workingStyle === "出社" ? ":shussha:" : workingStyle === "リモート" ? ":remote:" : "";
    if (restStartTime === undefined || restEndTime === undefined) {
      return `• ${dayOfWeek}: ${emojiWorkingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`;
    } else {
      return `• ${dayOfWeek}: ${emojiWorkingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")} (休憩: ${restStartTime}~${restEndTime})`;
    }
  });

  return `[追加]\n${messages.join("\n")}`;
};

const createMessageForModifyRecurringEvent = (
  beforeModificationInfos: Event[],
  afterModificationInfos: { title: string; dayOfWeek: DayOfWeek; startTime: Date; endTime: Date }[],
): string => {
  const beforeMessages = beforeModificationInfos.map(({ title, startTime, endTime }) => {
    const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(title);
    const emojiWorkingStyle = workingStyle === "出社" ? ":shussha:" : workingStyle === "リモート" ? ":remote:" : "";
    if (restStartTime === undefined || restEndTime === undefined) {
      return `${emojiWorkingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`;
    } else {
      return `${emojiWorkingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")} (休憩: ${restStartTime}~${restEndTime})`;
    }
  });
  const afterMessages = afterModificationInfos.map(({ title, startTime, endTime }) => {
    const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(title);
    const emojiWorkingStyle = workingStyle === "出社" ? ":shussha:" : workingStyle === "リモート" ? ":remote:" : "";
    if (restStartTime === undefined || restEndTime === undefined) {
      return `${emojiWorkingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`;
    } else {
      return `${emojiWorkingStyle} ${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")} (休憩: ${restStartTime}~${restEndTime})`;
    }
  });
  const messages = beforeMessages.map((message, index) => {
    return `• ${afterModificationInfos[index].dayOfWeek}: ${message} → ${afterMessages[index]}`;
  });
  return `[変更]\n${messages.join("\n")}`;
};

const createMessageForDeleteRecurringEvent = (deleteEvens: Event[], deletionInfos: DayOfWeek[]): string => {
  const message = deleteEvens.map(({ title, startTime, endTime }, index) => {
    const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(title);
    const emojiWorkingStyle = workingStyle === "出社" ? ":shussha:" : workingStyle === "リモート" ? ":remote:" : "";
    if (restStartTime === undefined || restEndTime === undefined) {
      return `• ${deletionInfos[index]}: ${emojiWorkingStyle}${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")}`;
    } else {
      return `• ${deletionInfos[index]}: ${emojiWorkingStyle}${format(startTime, "HH:mm")}~${format(endTime, "HH:mm")} (休憩: ${restStartTime}~${restEndTime})`;
    }
  });

  return `[消去]\n${message.join("\n")}`;
};

const createMessageForRecurringEvent = (
  { job, lastName }: PartTimerProfile,
  after: Date,
  registerEventStrings: string,
  modifyEventStrings: string,
  deleteEventStrings: string,
  comment: string,
): string => {
  const message = [
    `🔄${job}${lastName}さんが${format(after, "yyyy/MM/dd")}以降の固定シフトを変更しました🔄`,
    registerEventStrings,
    modifyEventStrings,
    deleteEventStrings,
  ].join("\n");

  return `${message}\n---\nコメント: ${comment}`;
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
  const date = format(eventInfo.startTime, "MM/dd");
  const { workingStyle, restStartTime, restEndTime } = getEventInfoFromTitle(eventInfo.title);
  const emojiWorkingStyle = workingStyle === "出社" ? ":shussha:" : workingStyle === "リモート" ? ":remote:" : "";
  const startTime = format(eventInfo.startTime, "HH:mm");
  const endTime = format(eventInfo.endTime, "HH:mm");
  if (restStartTime === undefined || restEndTime === undefined)
    return `• ${date}: ${emojiWorkingStyle} ${startTime}~${endTime}`;
  else return `• ${date}: ${emojiWorkingStyle} ${startTime}~${endTime} (休憩: ${restStartTime}~${restEndTime})`;
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
