import { z } from "zod";

import { Comment, DateOrEmptyString } from "./common.schema";
import { mergeTimeToDate } from "./date-utils";

const RegistrationSheetRow = z
  .object({
    date: z.date(),
    startTime: z.date(),
    endTime: z.date(),
    restStartTime: DateOrEmptyString,
    restEndTime: DateOrEmptyString,
    workingStyle: z.literal("出社").or(z.literal("リモート")),
  })
  .transform((row) => ({
    ...row,
    startTime: mergeTimeToDate(row.date, row.startTime),
    endTime: mergeTimeToDate(row.date, row.endTime),
  }))
  .refine((data) => (data.restStartTime && data.restEndTime ? data.restStartTime < data.restEndTime : true), {
    message: "休憩時間の開始時間が終了時間よりも前になるようにしてください",
  })
  .refine((data) => data.startTime > new Date() || data.endTime > new Date(), {
    message: "過去の時間にシフト変更はできません",
  });
type RegistrationSheetRow = z.infer<typeof RegistrationSheetRow>;

const RegistrationSheetValues = z.object({
  comment: Comment,
  registrationRows: z.array(RegistrationSheetRow),
});

export const insertRegistrationSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet: GoogleAppsScript.Spreadsheet.Sheet;
  try {
    sheet = spreadsheet.insertSheet("単発シフト登録", 0);
  } catch {
    throw new Error("既存の「単発シフト登録」シートを使用してください");
  }
  sheet.addDeveloperMetadata("part-timer-shift-manager-registration");
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
): {
  comment: Comment;
  registrationRows: RegistrationSheetRow[];
} => {
  const sheetRows = getRegistrationRows(sheet);
  const comment = sheet.getRange("A2").getValue();
  return RegistrationSheetValues.parse({ comment, registrationRows: sheetRows });
};

const getRegistrationRows = (sheet: GoogleAppsScript.Spreadsheet.Sheet): RegistrationSheetRow[] => {
  const sheetValues = sheet
    .getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn())
    .getValues()
    .map((row) =>
      RegistrationSheetRow.parse({
        date: row[0],
        startTime: row[1],
        endTime: row[2],
        restStartTime: row[3],
        restEndTime: row[4],
        workingStyle: row[5],
      }),
    );
  return sheetValues;
};
