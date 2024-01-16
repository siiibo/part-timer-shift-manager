import { set } from "date-fns";

import { PartTimerProfile } from "./JobSheet";
import { createTitleFromEventInfo } from "./shift-changer";
import { EventInfo } from "./shift-changer-api";

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

export const setValuesModificationAndDeletionSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
  const description1 = "コメント欄 (下の色付きセルに記入してください)";
  sheet.getRange("A1").setValue(description1).setFontWeight("bold");
  const commentCell = sheet.getRange("A2");
  commentCell.setBackground("#f0f8ff");

  const description2 = "本日以降の日付を下の色付きセルに記入してください。一ヶ月後までの予定が表示されます。";
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
export const getModificationAndDeletionSheetValues = (
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
      //NOTE: シート上の時間情報はDate型で取得される
      const date = new Date(row[1]);
      const startTime = row[2] as Date;
      const endTime = row[3] as Date;
      const newDate = row[4];
      const newStartTime = row[5];
      const newEndTime = row[6];
      return {
        title: row[0] as string,
        date: date,
        startTime: startTime,
        endTime: endTime,
        newDate: newDate,
        newStartTime: newStartTime,
        newEndTime: newEndTime,
        newRestStartTime: row[7] === "" ? undefined : row[7],
        newRestEndTime: row[8] === "" ? undefined : row[8],
        newWorkingStyle: row[9] as string,
        deletionFlag: row[10] as boolean,
      };
    });

  return sheetValues;
};

export const getModificationInfos = (
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
      const date = row.date;
      const startTime = set(date, {
        hours: row.startTime.getHours(),
        minutes: row.startTime.getMinutes(),
      });
      const endTime = set(date, { hours: row.endTime.getHours(), minutes: row.endTime.getMinutes() });
      const newDate = new Date(row.newDate);
      const newStartTime = set(newDate, {
        hours: row.newStartTime.getHours(),
        minutes: row.newStartTime.getMinutes(),
      });
      const newEndTime = set(newDate, {
        hours: row.newEndTime.getHours(),
        minutes: row.newEndTime.getMinutes(),
      });
      const newWorkingStyle = row.newWorkingStyle;
      if (newWorkingStyle === undefined) throw new Error("new working style is not defined");
      if (row.newRestStartTime === undefined || row.newRestEndTime === undefined) {
        const newTitle = createTitleFromEventInfo({ workingStyle: newWorkingStyle }, partTimerProfile);
        return {
          previousEventInfo: { title, date, startTime, endTime },
          newEventInfo: { title: newTitle, date: newDate, startTime: newStartTime, endTime: newEndTime },
        };
      } else {
        const newTitle = createTitleFromEventInfo(
          { restStartTime: row.newRestStartTime, restEndTime: row.newRestEndTime, workingStyle: newWorkingStyle },
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
export const getDeletionInfos = (
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
      const date = row.date;
      const startTime = set(date, {
        hours: row.startTime.getHours(),
        minutes: row.startTime.getMinutes(),
      });
      const endTime = set(date, { hours: row.endTime.getHours(), minutes: row.endTime.getMinutes() });
      return { title, date, startTime, endTime };
    });

  return deletionInfos;
};
