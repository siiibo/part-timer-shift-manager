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
const setValuesModificationAndDeletionSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
  const description1 = "コメント欄 (下の色付きセルに記入してください)";
  sheet.getRange("A1").setValue(description1).setFontWeight("bold");
  const commentCell = sheet.getRange("A2");
  commentCell.setBackground("#f0f8ff");

  const description2 = "本日以降の日付を下の色付きセルに記入してください。一週間後までの予定が表示されます。";
  sheet.getRange("A4").setValue(description2).setFontWeight("bold");
  const dateCell = sheet.getRange("A5");
  dateCell.setBackground("#f0f8ff");

  const description3 = "【予定一覧】";
  sheet.getRange("A7").setValue(description3).setFontWeight("bold");

  const description4 = "【変更】変更後の予定を記入してください ";
  sheet.getRange("E7").setValue(description4).setFontWeight("bold");

  const description5 = "【削除】削除したい予定を選択してください";
  sheet.getRange("K7").setValue(description5).setFontWeight("bold");

  const header = [
    "イベント名",
    "日付",
    "開始時刻",
    "終了時刻",
    "日付",
    "開始時刻",
    "終了時刻",
    "休憩開始時刻",
    "休憩終了時刻",
    "勤務形態",
    "削除対象",
  ];
  sheet.getRange(8, 1, 1, header.length).setValues([header]).setFontWeight("bold");

  const dateCells = sheet.getRange("E9:E1000");
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDateOnOrAfter(new Date())
    .setAllowInvalid(false)
    .setHelpText("本日以降の日付を入力してください。")
    .build();
  dateCell.setDataValidation(dateRule);
  dateCells.setDataValidation(dateRule);
  const timeCells = sheet.getRange("F9:I1000");
  const timeRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied("=ISDATE(F9)")
    .setAllowInvalid(false)
    .setHelpText('時刻を"◯◯:◯◯"の形式で入力してください。\n【例】 9:00')
    .build();
  timeCells.setDataValidation(timeRule);
  const workingStyleCells = sheet.getRange("J9:J1000");
  const workingStyleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リモート", "出社"], true)
    .setAllowInvalid(false)
    .setHelpText("リモート/出社 を選択してください。")
    .build();
  workingStyleCells.setDataValidation(workingStyleRule);
  const checkboxCells = sheet.getRange("K9:K1000");
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .setHelpText("チェックボックス以外の入力形式は認められません。")
    .build();
  checkboxCells.setDataValidation(checkboxRule);

  sheet.setColumnWidth(1, 370);
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

const getModificationAndDeletionSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): {
  title: string;
  date: Date;
  startTime: Date;
  endTime: Date;
  newDate: Date;
  newStartTime: Date;
  newEndTime: Date;
  newRestStartTime?: Date;
  newRestEndTime?: Date;
  newWorkingStyle: string;
  deletionFlag: boolean;
}[] => {
  const sheetValues = sheet
    .getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn())
    .getValues()
    .map((row) => {
      return {
        title: row[0] as string,
        date: row[1] as Date,
        startTime: row[2] as Date,
        endTime: row[3] as Date,
        newDate: row[4] as Date,
        newStartTime: row[5] as Date,
        newEndTime: row[6] as Date,
        newRestStartTime: row[7] === "" ? undefined : row[7],
        newRestEndTime: row[8] === "" ? undefined : row[8],
        newWorkingStyle: row[9] as string,
        deletionFlag: row[10] as boolean,
      };
    });

  return sheetValues;
};


const getModificationInfos = (
  sheetValues: {
    title: string;
    date: Date;
    startTime: Date;
    endTime: Date;
    newDate: Date;
    newStartTime: Date;
    newEndTime: Date;
    newRestStartTime?: Date;
    newRestEndTime?: Date;
    newWorkingStyle?: string;
    deletionFlag: boolean;
  }[],
  partTimerProfile: PartTimerProfile
): {
  previousEventInfo: EventInfo;
  newEventInfo: EventInfo;
}[] => {
  const modificationInfos = sheetValues
    .filter((row) => !row.deletionFlag)
    .map((row) => {
      const title = row.title;
      const date = format(row.date, "yyyy-MM-dd");
      const startTime = format(row.startTime, "HH:mm");
      const endTime = format(row.endTime, "HH:mm");
      const newDate = format(row.newDate, "yyyy-MM-dd");
      const newStartTime = format(row.newStartTime, "HH:mm");
      const newEndTime = format(row.newEndTime, "HH:mm");
      const newWorkingStyle = row.newWorkingStyle;
      if (newWorkingStyle === undefined) throw new Error("new working style is not defined");
      if (row.newRestStartTime === undefined || row.newRestEndTime === undefined) {
        const newTitle = createTitleFromEventInfo({ workingStyle: newWorkingStyle }, partTimerProfile);
        return {
          previousEventInfo: { title, date, startTime, endTime },
          newEventInfo: { title: newTitle, date: newDate, startTime: newStartTime, endTime: newEndTime },
        };
      } else {
        const newRestStartTime =row.newRestStartTime;
        const newRestEndTime = row.newRestEndTime;
        const newTitle = createTitleFromEventInfo(
          { restStartTime: newRestStartTime, restEndTime: newRestEndTime, workingStyle: newWorkingStyle },
          partTimerProfile
        );
        return {
          previousEventInfo: { title, date, startTime, endTime },
          newEventInfo: { title: newTitle, date: newDate, startTime: newStartTime, endTime: newEndTime },
        };
      }
    });

  return modificationInfos;
};
const getDeletionInfos = (
    sheetValues: {
      title: string;
      date: Date;
      startTime: Date;
      endTime: Date;
      newDate: Date;
      newStartTime: Date;
      newEndTime: Date;
      newRestStartTime?: Date;
      newRestEndTime?: Date;
      newWorkingStyle: string;
      deletionFlag: boolean;
    }[]
  ): EventInfo[] => {
    const deletionInfos = sheetValues
      .filter((row) => row.deletionFlag)
      .map((row) => {
        const title = row.title;
        const date = format(row.date, "yyyy-MM-dd");
        const startTime = format(row.startTime, "HH:mm");
        const endTime = format(row.endTime, "HH:mm");
        return { title, date, startTime, endTime };
      });
  
    return deletionInfos;
  };
const createModificationMessage = (
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
const createDeletionMessage = (deletionInfos: EventInfo[], partTimerProfile: PartTimerProfile): string | undefined => {
  const messages = deletionInfos.map(createMessageFromEventInfo);
  if (messages.length == 0) return;
  const { job, lastName } = partTimerProfile;
  const messageTitle = `${job}${lastName}さんの以下の予定が削除されました。`;
  return `${messageTitle}\n${messages.join("\n")}`;
};

