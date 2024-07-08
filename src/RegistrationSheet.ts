import { set } from "date-fns";
import { z } from "zod";

import { DateAfterNow } from "./common.schema";

const RegistrationRow = z
  .object({
    startTime: DateAfterNow,
    endTime: DateAfterNow,
    restStartTime: z.date().optional(),
    restEndTime: z.date().optional(),
    workingStyle: z.literal("出社").or(z.literal("リモート")),
  })
  .refine(
    (data) => {
      if (data.restStartTime && data.restEndTime) {
        return data.restStartTime < data.restEndTime;
      }
      return true;
    },
    {
      message: "休憩時間の開始時間が終了時間よりも前になるようにしてください",
    },
  );
type RegistrationRow = z.infer<typeof RegistrationRow>;

export const insertRegistrationSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`単発シフト登録`, 0);
  } catch {
    throw new Error("既存の「単発シフト登録」シートを使用してください");
  }
  sheet.addDeveloperMetadata(`part-timer-shift-manager-registration`);
  setValuesRegistrationSheet(sheet);
};

export const setValuesRegistrationSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
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

export const getRegistrationRows = (sheet: GoogleAppsScript.Spreadsheet.Sheet): RegistrationRow[] => {
  const registrationRows = sheet
    .getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn())
    .getValues()
    .map((eventInfo) => {
      //NOTE: セルの書式設定が日付になっている場合はDate型が渡ってくる
      const date = eventInfo[0];
      const startTimeDate = eventInfo[1];
      const startTime = set(date, {
        hours: startTimeDate.getHours(),
        minutes: startTimeDate.getMinutes(),
      });
      const endTimeDate = eventInfo[2];
      const endTime = set(date, { hours: endTimeDate.getHours(), minutes: endTimeDate.getMinutes() });
      const restStartTime = eventInfo[3] === "" ? undefined : eventInfo[3];
      const restEndTime = eventInfo[4] === "" ? undefined : eventInfo[4];
      const workingStyle = eventInfo[5];
      return RegistrationRow.parse({
        startTime,
        endTime,
        restStartTime,
        restEndTime,
        workingStyle,
      });
    });
  return registrationRows;
};
