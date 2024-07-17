import { set } from "date-fns";
import { z } from "zod";

import { Comment, DateAfterNow, DateOrEmptyString } from "./common.schema";

const RegistrationSheetRow = z.object({
  comment: Comment,
  date: z.date(),
  startTimeDate: z.date(),
  endTimeDate: z.date(),
  restStartTime: DateOrEmptyString,
  restEndTime: DateOrEmptyString,
  workingStyle: z.literal("出社").or(z.literal("リモート")),
});
type RegistrationSheetRow = z.infer<typeof RegistrationSheetRow>;

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

export const getRegistrationSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): { comment: string; sheetValues: RegistrationRow[] } => {
  const sheetRows = getRegistrationRows(sheet);
  const comment = sheetRows[0].comment;
  const sheetValues = sheetRows.map(
    ({ date, startTimeDate, endTimeDate, restStartTime, restEndTime, workingStyle }) => {
      const startTime = mergeTimeToDate(date, startTimeDate);
      const endTime = mergeTimeToDate(date, endTimeDate);
      return RegistrationRow.parse({
        startTime,
        endTime,
        restStartTime,
        restEndTime,
        workingStyle,
      });
    },
  );
  return { comment, sheetValues };
};

const getRegistrationRows = (sheet: GoogleAppsScript.Spreadsheet.Sheet): RegistrationSheetRow[] => {
  const comment = sheet.getRange("A2").getValue();
  const sheetValues = sheet
    .getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn())
    .getValues()
    .map((eventInfo) => {
      const date = eventInfo[0];
      const startTimeDate = eventInfo[1];
      const endTimeDate = eventInfo[2];
      const restStartTime = eventInfo[3];
      const restEndTime = eventInfo[4];
      const workingStyle = eventInfo[5];
      return RegistrationSheetRow.parse({
        comment,
        date,
        startTimeDate,
        endTimeDate,
        restStartTime,
        restEndTime,
        workingStyle,
      });
    });
  return sheetValues;
};

//NOTE: Googleスプレッドシートでは時間のみの入力がDate型として取得される際、日付部分はデフォルトで1899/12/30となるため適切な日付情報に更新する必要がある
const mergeTimeToDate = (date: Date, time: Date): Date => {
  return set(date, { hours: time.getHours(), minutes: time.getMinutes() });
};
