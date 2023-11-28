import { format } from "date-fns";

import { getConfig } from "./config";
import { EventInfo } from "./shift-changer-api";
import {
  createMessageFromEventInfo,
  createTitleFromEventInfo,
  getPartTimerProfile,
  getSheet,
  getSlackClient,
  postMessageToSlackChannel,
} from "./utils";

type SheetType = "registration" | "modificationAndDeletion";
type OperationType = "registration" | "modificationAndDeletion" | "showEvents";
type PartTimerProfile = {
  job: string;
  lastName: string;
  email: string;
  managerEmails: string[];
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
const setValuesRegistrationSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
  const description1 = "コメント欄 (下の色付きセルに記入してください)";
  sheet.getRange("A1").setValue(description1).setFontWeight("bold");
  const commentCell = sheet.getRange("A2");
  commentCell.setBackground("#f0f8ff");

  const header = ["日付", "開始時刻", "終了時刻", "休憩開始時刻", "休憩終了時刻", "勤務形態"];
  sheet.getRange(4, 1, 1, header.length).setValues([header]).setFontWeight("bold");

  const workingStyleCells = sheet.getRange("F5:F1000");
  const workingStyleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リモート", "出社"], true)
    .setAllowInvalid(false)
    .setHelpText("リモート/出社 を選択してください。")
    .build();
  workingStyleCells.setDataValidation(workingStyleRule);
  const dateCells = sheet.getRange("A5:A1000");
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDateOnOrAfter(new Date())
    .setAllowInvalid(false)
    .setHelpText("本日以降の日付を入力してください。")
    .build();
  dateCells.setDataValidation(dateRule);
  const timeCells = sheet.getRange("B5:E1000");
  const timeRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied("=ISDATE(B5)")
    .setHelpText('時刻を"◯◯:◯◯"の形式で入力してください。')
    .build();
  timeCells.setDataValidation(timeRule);
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

const getRegistrationInfos = (
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    partTimerProfile: PartTimerProfile
  ): EventInfo[] => {
    const registrationInfos = sheet
      .getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn())
      .getValues()
      .map((eventInfo) => {
        const date = format(eventInfo[0] as Date, "yyyy-MM-dd");
        const startTime = format(eventInfo[1] as Date, "HH:mm");
        const endTime = format(eventInfo[2] as Date, "HH:mm");
        const workingStyle = eventInfo[5] as string;
        if (workingStyle === "") throw new Error("working style is not defined");
        if (eventInfo[3] === "" || eventInfo[4] === "") {
          const title = createTitleFromEventInfo({ workingStyle }, partTimerProfile);
          return { title, date, startTime, endTime };
        } else {
          const restStartTime = eventInfo[3];
          const restEndTime = eventInfo[4];
          const title = createTitleFromEventInfo({ restStartTime, restEndTime, workingStyle }, partTimerProfile);
          return { title, date, startTime, endTime };
        }
      });
    return registrationInfos;
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
