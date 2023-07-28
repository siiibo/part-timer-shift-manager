import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { format } from "date-fns";

import { getConfig } from "./config";
import { EventInfo, shiftChanger } from "./shift-changer-api";

type SheetType = "registration" | "modificationAndDeletion";
type OperationType = "registration" | "modificationAndDeletion" | "showEvents";

export const doPost = (e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput => {
  if (e.parameter.apiId === "shift-changer") {
    const response = shiftChanger(e) ?? "";
    return ContentService.createTextOutput(response).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput("undefined");
};

export const onOpen = () => {
  const ui = SpreadsheetApp.getUi();

  ui.createAddonMenu()
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
  const today = format(new Date(), "yyyy-MM-dd");
  const sheet = spreadsheet.insertSheet(`${today}-登録`, 0);
  sheet.addDeveloperMetadata(`${today}-registration`);

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

export const insertModificationAndDeletionSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const today = format(new Date(), "yyyy-MM-dd");
  const sheet = spreadsheet.insertSheet(`${today}-変更・削除`, 0);
  sheet.addDeveloperMetadata(`${today}-modificationAndDeletion`);

  const description1 = "コメント欄 (下の色付きセルに記入してください)";
  sheet.getRange("A1").setValue(description1).setFontWeight("bold");
  const commentCell = sheet.getRange("A2");
  commentCell.setBackground("#f0f8ff");

  const description2 = "本日以降の日付を入力してください。指定した日付から一週間後までの予定が表示されます。";
  sheet.getRange("A4").setValue(description2).setFontWeight("bold");

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

  const dateCell = sheet.getRange("A5");
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
export const callRegistration = () => {
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const { SLACK_ACCESS_TOKEN } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  const slackMemberProfiles = getSlackMemberProfiles(client);

  const sheetType: SheetType = "registration";
  const sheet = getSheet(sheetType, spreadsheetUrl);
  const operationType: OperationType = "registration";
  const comment = sheet.getRange("A2").getValue();
  const registrationInfos = getRegistrationInfos(sheet, userEmail, slackMemberProfiles);

  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    registrationInfos: JSON.stringify(registrationInfos),
  };
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
  };
  const { API_URL, SLACK_CHANNEL_TO_POST } = getConfig();
  UrlFetchApp.fetch(API_URL, options);
  const messageToNotify = createRegistrationMessage(registrationInfos, comment);
  postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, messageToNotify, userEmail);
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
  newRestStartTime: Date | string;
  newRestEndTime: Date | string;
  newWorkingStyle: string;
  deletionFlag: boolean;
}[] => {
  const sheetValues = sheet
    .getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn())
    .getValues()
    .map((row) => {
      if (row[7] === "" || row[8] === "") {
        return {
          title: row[0] as string,
          date: row[1] as Date,
          startTime: row[2] as Date,
          endTime: row[3] as Date,
          newDate: row[4] as Date,
          newStartTime: row[5] as Date,
          newEndTime: row[6] as Date,
          newRestStartTime: row[7] as string,
          newRestEndTime: row[8] as string,
          newWorkingStyle: row[9] as string,
          deletionFlag: row[10] as boolean,
        };
      } else {
        return {
          title: row[0] as string,
          date: row[1] as Date,
          startTime: row[2] as Date,
          endTime: row[3] as Date,
          newDate: row[4] as Date,
          newStartTime: row[5] as Date,
          newEndTime: row[6] as Date,
          newRestStartTime: row[7] as Date,
          newRestEndTime: row[8] as Date,
          newWorkingStyle: row[9] as string,
          deletionFlag: row[10] as boolean,
        };
      }
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
    newRestStartTime: Date | string;
    newRestEndTime: Date | string;
    newWorkingStyle: string;
    deletionFlag: boolean;
  }[],
  userEmail: string,
  slackMemberProfiles: {
    name: string;
    email: string;
  }[]
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
      if (row.newRestStartTime === "" || row.newRestEndTime === "") {
        const newRestStartTime = row.newRestStartTime as string;
        const newRestEndTime = row.newRestEndTime as string;
        const newWorkingStyle = row.newWorkingStyle;
        if (newWorkingStyle === "") throw new Error("new working style is not defined");
        const newTitle = createTitleFromEventInfo(
          { restStartTime: newRestStartTime, restEndTime: newRestEndTime, workingStyle: newWorkingStyle },
          userEmail,
          slackMemberProfiles
        );
        return {
          previousEventInfo: { title, date, startTime, endTime },
          newEventInfo: { title: newTitle, date: newDate, startTime: newStartTime, endTime: newEndTime },
        };
      } else {
        const newRestStartTime = format(row.newRestStartTime as Date, "HH:mm");
        const newRestEndTime = format(row.newRestEndTime as Date, "HH:mm");
        const newWorkingStyle = row.newWorkingStyle;
        const newTitle = createTitleFromEventInfo(
          { restStartTime: newRestStartTime, restEndTime: newRestEndTime, workingStyle: newWorkingStyle },
          userEmail,
          slackMemberProfiles
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
    newRestStartTime: Date | string;
    newRestEndTime: Date | string;
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
export const callModificationAndDeletion = () => {
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const { SLACK_ACCESS_TOKEN } = getConfig();
  const client = getSlackClient(SLACK_ACCESS_TOKEN);
  const slackMemberProfiles = getSlackMemberProfiles(client);
  const sheetType: SheetType = "modificationAndDeletion";
  const sheet = getSheet(sheetType, spreadsheetUrl);
  const comment = sheet.getRange("A2").getValue();
  const operationType: OperationType = "modificationAndDeletion";
  const sheetValues = getModificationAndDeletionSheetValues(sheet);
  const valuesForOperation = sheetValues.filter((row) => row.deletionFlag || row.newDate);
  const modificationInfos = getModificationInfos(valuesForOperation, userEmail, slackMemberProfiles);
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
  };
  const { API_URL, SLACK_CHANNEL_TO_POST } = getConfig();
  UrlFetchApp.fetch(API_URL, options);

  const modificationMessageToNotify = createModificationMessage(modificationInfos, comment);
  if (modificationMessageToNotify)
    postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, modificationMessageToNotify, userEmail);

  const deletionMessageToNotify = createDeletionMessage(deletionInfos, comment);
  if (deletionMessageToNotify)
    postMessageToSlackChannel(client, SLACK_CHANNEL_TO_POST, deletionMessageToNotify, userEmail);
};

export const callShowEvents = () => {
  const userEmail = Session.getActiveUser().getEmail();
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheetType: SheetType = "modificationAndDeletion";
  const sheet = getSheet(sheetType, spreadsheetUrl);
  const operationType: OperationType = "showEvents";
  const startDate: Date = sheet.getRange("A5").getValue();

  sheet.getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn()).clearContent();

  const payload = {
    apiId: "shift-changer",
    operationType: operationType,
    userEmail: userEmail,
    startDate: startDate,
  };
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    payload: payload,
  };
  const { API_URL } = getConfig();
  const response = UrlFetchApp.fetch(API_URL, options);
  if (!response.getContentText()) return;

  const eventInfos: EventInfo[] = JSON.parse(response.getContentText());
  if (eventInfos.length === 0) throw new Error("no events");

  const moldedEventInfos = eventInfos.map(({ title, date, startTime, endTime }) => {
    return [title, date, startTime, endTime];
  });

  sheet.getRange(9, 1, moldedEventInfos.length, moldedEventInfos[0].length).setValues(moldedEventInfos);
};

const getSheet = (sheetType: SheetType, spreadsheetUrl: string): GoogleAppsScript.Spreadsheet.Sheet => {
  const today = format(new Date(), "yyyy-MM-dd");
  const sheet = SpreadsheetApp.openByUrl(spreadsheetUrl)
    .getSheets()
    .find((sheet) => sheet.getDeveloperMetadata().some((metaData) => metaData.getKey() === `${today}-${sheetType}`));

  if (!sheet) throw new Error("SHEET is not defined");

  return sheet;
};

const getRegistrationInfos = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  userEmail: string,
  slackMemberProfiles: { name: string; email: string }[]
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
        const restStartTime = eventInfo[3] as string;
        const restEndTime = eventInfo[4] as string;
        const title = createTitleFromEventInfo(
          { restStartTime, restEndTime, workingStyle },
          userEmail,
          slackMemberProfiles
        );
        return { title, date, startTime, endTime };
      } else {
        const restStartTime = format(eventInfo[3] as Date, "HH:mm");
        const restEndTime = format(eventInfo[4] as Date, "HH:mm");
        const title = createTitleFromEventInfo(
          { restStartTime, restEndTime, workingStyle },
          userEmail,
          slackMemberProfiles
        );
        return { title, date, startTime, endTime };
      }
    });
  return registrationInfos;
};

const createTitleFromEventInfo = (
  eventInfo: {
    restStartTime: string;
    restEndTime: string;
    workingStyle: string;
  },
  userEmail: string,
  slackMemberProfiles: {
    name: string;
    email: string;
  }[]
): string => {
  const name = getNameFromEmail(userEmail, slackMemberProfiles);
  const job = getJob(userEmail);

  const restStartTime = eventInfo.restStartTime;
  const restEndTime = eventInfo.restEndTime;
  const workingStyle = eventInfo.workingStyle;

  if (restStartTime === "" || restEndTime === "") {
    const title = `【${workingStyle}】${job}: ${name}さん`;
    return title;
  } else {
    const title = `【${workingStyle}】${job}: ${name}さん (休憩: ${restStartTime}~${restEndTime})`;
    return title;
  }
};

const getNameFromEmail = (email: string, slackMemberProfiles: { name: string; email: string }[]): string => {
  const slackMember = slackMemberProfiles.find((slackMemberProfile) => slackMemberProfile.email === email);
  if (!slackMember) throw new Error("The email is non-slack member");
  return slackMember.name;
};

const getSlackMemberProfiles = (client: SlackClient): { name: string; email: string }[] => {
  const slackMembers = client.users.list().members ?? [];

  const siiiboSlackMembers = slackMembers.filter(
    (slackMember) =>
      !slackMember.deleted &&
      !slackMember.is_bot &&
      slackMember.id !== "USLACKBOT" &&
      slackMember.profile?.email?.includes("siiibo.com")
  );

  const slackMemberProfiles = siiiboSlackMembers
    .map((slackMember) => {
      return {
        name: slackMember.profile?.real_name,
        email: slackMember.profile?.email,
      };
    })
    .filter((s): s is { name: string; email: string } => s.name !== "" || s.email !== "");
  return slackMemberProfiles;
};

const getSlackClient = (slackToken: string): SlackClient => {
  return new SlackClient(slackToken);
};

const getJob = (userEmail: string): string => {
  const { JOB_SHEET_URL } = getConfig();
  const sheet = SpreadsheetApp.openByUrl(JOB_SHEET_URL).getSheetByName("シート1");
  if (!sheet) throw new Error("SHEET is not defined");
  const partTimerInfos = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const partTimerInfo = partTimerInfos.find((partTimerInfo) => {
    const email = partTimerInfo[2] as string;
    return email === userEmail;
  });
  if (partTimerInfo === undefined) throw new Error("no part timer information for the email");

  const job = partTimerInfo[0] as string;
  return job;
};

const createMessageFromEventInfo = (eventInfo: EventInfo) => {
  const formattedDate = format(new Date(eventInfo.date), "MM/dd");
  return `${eventInfo.title}: ${formattedDate} ${eventInfo.startTime}~${eventInfo.endTime}`;
};

const createRegistrationMessage = (registrationInfos: EventInfo[], comment: string): string => {
  const messages = registrationInfos.map(createMessageFromEventInfo);
  const messageTitle = "以下の予定が追加されました。";
  return comment
    ? `${messageTitle}\n${messages.join("\n")}\n\nコメント: ${comment}`
    : `${messageTitle}\n${messages.join("\n")}`;
};

const createDeletionMessage = (deletionInfos: EventInfo[], comment: string): string | undefined => {
  const messages = deletionInfos.map(createMessageFromEventInfo);
  if (messages.length == 0) return;
  const messageTitle = "以下の予定が削除されました。";
  return comment
    ? `${messageTitle}\n${messages.join("\n")}\n\nコメント: ${comment}`
    : `${messageTitle}\n${messages.join("\n")}`;
};

const createModificationMessage = (
  modificationInfos: {
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  }[],
  comment: string
): string | undefined => {
  const messages = modificationInfos.map(({ previousEventInfo, newEventInfo }) => {
    return `${createMessageFromEventInfo(previousEventInfo)}\n\
    → ${createMessageFromEventInfo(newEventInfo)}`;
  });
  if (messages.length == 0) return;
  const messageTitle = "以下の予定が変更されました。";
  return comment
    ? `${messageTitle}\n${messages.join("\n")}\n\nコメント: ${comment}`
    : `${messageTitle}\n${messages.join("\n")}`;
};

const getManagerEmails = (userEmail: string): string[] => {
  const { JOB_SHEET_URL } = getConfig();
  const sheet = SpreadsheetApp.openByUrl(JOB_SHEET_URL).getSheetByName("シート1");
  if (!sheet) throw new Error("SHEET is not defined");
  const partTimerInfos = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const partTimerInfo = partTimerInfos.find((partTimerInfo) => {
    const email = partTimerInfo[2] as string;
    return email === userEmail;
  });
  if (partTimerInfo === undefined) throw new Error("no part timer information for the email");
  const managerEmail = partTimerInfo[3] as string;
  const managerEmails = managerEmail.replaceAll(/\s/g, "").split(",");
  return managerEmails;
};

const getManagerSlackIds = (managerEmails: string[], client: SlackClient): string[] => {
  const slackMembers = client.users.list().members ?? [];

  const managerSlackIds = managerEmails
    .map((email) => {
      const member = slackMembers.find((slackMember) => {
        return slackMember.profile?.email === email;
      });
      if (member === undefined) throw new Error("The email is not in the slack members");
      return member.id;
    })
    .filter((id): id is string => id !== undefined);

  return managerSlackIds;
};

const slackIdToMention = (slackId: string) => `<@${slackId}>`;

const postMessageToSlackChannel = (
  client: SlackClient,
  slackChannelToPost: string,
  messageToNotify: string,
  userEmail: string
) => {
  const { HR_MANAGER_SLACK_ID } = getConfig();
  const managerEmails = getManagerEmails(userEmail);
  const managerSlackIds = getManagerSlackIds(managerEmails, client);
  const mentionMessageToManagers = [HR_MANAGER_SLACK_ID, ...managerSlackIds].map(slackIdToMention).join(" ");
  client.chat.postMessage({
    channel: slackChannelToPost,
    text: `${mentionMessageToManagers}\n${messageToNotify}`,
  });
};
