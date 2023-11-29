import { getConfig } from "./config";
import {
  getDeletionInfos,
  getModificationAndDeletionSheetValues,
  getModificationInfos,
  setValuesModificationAndDeletionSheet,
} from "./ModificationAndDeletionSheet";
import {
  createDeletionMessage,
  createModificationMessage,
  getSlackClient,
  postMessageToSlackChannel,
} from "./Postslack";
import { createMessageFromEventInfo } from "./Postslack";
import { getRegistrationInfos, setValuesRegistrationSheet } from "./RegistrationSheet";
import { shiftChanger } from "./shift-changer-api";
import { EventInfo } from "./shift-changer-api";

type SheetType = "registration" | "modificationAndDeletion";
type OperationType = "registration" | "modificationAndDeletion" | "showEvents";
type PartTimerProfile = {
  job: string;
  lastName: string;
  email: string;
  managerEmails: string[];
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
        .addItem("提出", callRegistration.name)
    )
    .addSubMenu(
      ui
        .createMenu("変更・削除")
        .addItem("シートの追加", insertModificationAndDeletionSheet.name)
        .addSeparator()
        .addItem("予定を表示", callShowEvents.name)
        .addItem("提出", callModificationAndDeletion.name)
    )
    .addToUi();
};

export const insertRegistrationSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`登録`, 0);
  } catch {
    throw new Error("既存の「登録」シートを使用してください");
  }
  sheet.addDeveloperMetadata(`part-timer-shift-manager-registration`);
  setValuesRegistrationSheet(sheet);
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
  const registrationInfos = getRegistrationInfos(sheet, partTimerProfile);

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
  partTimerProfile: PartTimerProfile
): string => {
  const messages = registrationInfos.map(createMessageFromEventInfo);
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が追加されました。`;
  return comment
    ? `${messageTitle}\n${messages.join("\n")}\n\nコメント: ${comment}`
    : `${messageTitle}\n${messages.join("\n")}`;
};

export const insertModificationAndDeletionSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`変更・削除`, 0);
  } catch {
    throw new Error("既存の「変更・削除」シートを使用してください");
  }
  sheet.addDeveloperMetadata(`part-timer-shift-manager-modificationAndDeletion`);
  setValuesModificationAndDeletionSheet(sheet);
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

  const eventInfos: EventInfo[] = JSON.parse(response.getContentText());
  if (eventInfos.length === 0) throw new Error("no events");

  const moldedEventInfos = eventInfos.map(({ title, date, startTime, endTime }) => {
    return [title, date, startTime, endTime];
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
  const sheetValues = getModificationAndDeletionSheetValues(sheet);
  const valuesForOperation = sheetValues.filter((row) => row.deletionFlag || row.newDate);
  const modificationInfos = getModificationInfos(valuesForOperation, partTimerProfile);
  const deletionInfos = getDeletionInfos(valuesForOperation);

  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    modificationInfos: JSON.stringify(modificationInfos),
    deletionInfos: JSON.stringify(deletionInfos),
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

  const modificationAndDeletionMessageToNotify = [
    createModificationMessage(modificationInfos, partTimerProfile),
    createDeletionMessage(deletionInfos, partTimerProfile),
    comment ? `コメント: ${comment}` : undefined,
  ]
    .filter(Boolean)
    .join("\n---\n");

  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, modificationAndDeletionMessageToNotify, partTimerProfile);
  sheet.clear();
  SpreadsheetApp.flush();
  setValuesModificationAndDeletionSheet(sheet);
};

const getPartTimerProfile = (userEmail: string): PartTimerProfile => {
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
const getSheet = (sheetType: SheetType, spreadsheetUrl: string): GoogleAppsScript.Spreadsheet.Sheet => {
  const sheet = SpreadsheetApp.openByUrl(spreadsheetUrl)
    .getSheets()
    .find((sheet) =>
      sheet.getDeveloperMetadata().some((metaData) => metaData.getKey() === `part-timer-shift-manager-${sheetType}`)
    );

  if (!sheet) throw new Error("SHEET is not defined");

  return sheet;
};
