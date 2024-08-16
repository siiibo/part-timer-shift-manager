import { z } from "zod";

import {
  Comment,
  DateAfterNow,
  DateOrEmptyString,
  DayOfWeek,
  DayOfWeekOrEmptyString,
  WorkingStyleOrEmptyString,
} from "./common.schema";

const OperationString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("追加").or(z.literal("時間変更")).or(z.literal("消去")).optional(),
);

const RecurringEventSheetRow = z
  .object({
    operation: OperationString,
    dayOfWeek: DayOfWeekOrEmptyString,
    startTime: DateOrEmptyString,
    endTime: DateOrEmptyString,
    restStartTime: DateOrEmptyString,
    restEndTime: DateOrEmptyString,
    workingStyle: WorkingStyleOrEmptyString,
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
type RecurringEventSheetRow = z.infer<typeof RecurringEventSheetRow>;

const RegisterRecurringEventRow = z.object({
  type: z.literal("registerRecurringEvent"),
  dayOfWeek: DayOfWeek,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("出社").or(z.literal("リモート")),
});
type RegisterRecurringEventRow = z.infer<typeof RegisterRecurringEventRow>;

const ModifyRecurringEventRow = z.object({
  type: z.literal("modifyRecurringEvent"),
  dayOfWeek: DayOfWeek,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("出社").or(z.literal("リモート")),
});
type ModifyRecurringEventRow = z.infer<typeof ModifyRecurringEventRow>;

const DeleteRecurringEventRow = z.object({
  type: z.literal("deleteRecurringEvent"),
  dayOfWeek: DayOfWeek,
});
type DeleteRecurringEventRow = z.infer<typeof DeleteRecurringEventRow>;

type NoOperationRow = {
  type: "no-operation";
};

type RecurringEventSheetValues = {
  after: DateAfterNow;
  comment: Comment;
  registrationRows: RegisterRecurringEventRow[];
  modificationRows: ModifyRecurringEventRow[];
  deletionRows: DeleteRecurringEventRow[];
};

export const insertRecurringEventSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`固定シフト`, 0);
  } catch {
    throw new Error("既存の「固定シフト」シートを使用してください");
  }
  sheet.addDeveloperMetadata(`part-timer-shift-manager-recurringEvent`);
  setValuesRecurringEventSheet(sheet);
};

export const setValuesRecurringEventSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
  const description1 = "コメント欄 (下の色付きセルに記入してください)";
  sheet.getRange("A1").setValue(description1).setFontWeight("bold");
  const commentCell = sheet.getRange("A2");
  commentCell.setBackground("#f0f8ff");

  const description2 = "本日以降の日付を下の色付きセルに記入してください。変更後のシフト開始日が設定できます";
  sheet.getRange("A4").setValue(description2).setFontWeight("bold");
  const dateCell = sheet.getRange("A5");
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDateOnOrAfter(new Date())
    .setAllowInvalid(false)
    .setHelpText("本日以降の日付を入力してください。")
    .build();
  dateCell.setBackground("#f0f8ff");
  dateCell.setDataValidation(dateRule);
  const description3 = "【固定シフト変更】";
  sheet.getRange("A7").setValue(description3).setFontWeight("bold");
  const header2 = ["操作", "曜日", "開始時間", "終了時刻", "休憩開始時刻", "休憩終了時刻", "勤務形態"];
  sheet.getRange(8, 1, 1, header2.length).setValues([header2]).setFontWeight("bold");
  const operationCells = sheet.getRange("A9:A13");
  const operationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["追加", "時間変更", "消去"], true)
    .setAllowInvalid(false)
    .setHelpText("登録, 時間変更, 削除 を選択してください。")
    .build();
  operationCells.setDataValidation(operationRule);
  sheet.getRange("B9:B13").setValues([["月曜日"], ["火曜日"], ["水曜日"], ["木曜日"], ["金曜日"]]);

  const workingStyleCells = sheet.getRange("G9:G13");
  const workingStyleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リモート", "出社"], true)
    .setAllowInvalid(false)
    .setHelpText("リモート/出社 を選択してください。")
    .build();
  workingStyleCells.setDataValidation(workingStyleRule);

  sheet.setColumnWidth(1, 370);
  sheet.setColumnWidth(2, 150);
};

export const getRecurringEventSheetValues = (sheet: GoogleAppsScript.Spreadsheet.Sheet): RecurringEventSheetValues => {
  const sheetRows = getRecurringEventSheetRows(sheet);
  console.log(sheetRows); //NOTE: シートの内容を確認するためのログ
  const after = DateAfterNow.parse(sheet.getRange("A5").getValue());
  const comment = Comment.parse(sheet.getRange("A2").getValue());
  const sheetValues = sheetRows.map((row) => {
    if (row.operation === "追加" && row.dayOfWeek && row.startTime && row.endTime) {
      return RegisterRecurringEventRow.parse({
        type: "registerRecurringEvent",
        dayOfWeek: row.dayOfWeek,
        startTime: row.startTime,
        endTime: row.endTime,
        restStartTime: row.restStartTime,
        restEndTime: row.restEndTime,
        workingStyle: row.workingStyle,
      });
    } else if (row.operation === "時間変更" && row.dayOfWeek && row.startTime && row.endTime) {
      return ModifyRecurringEventRow.parse({
        type: "modifyRecurringEvent",
        dayOfWeek: row.dayOfWeek,
        startTime: row.startTime,
        endTime: row.endTime,
        restStartTime: row.restStartTime,
        restEndTime: row.restEndTime,
        workingStyle: row.workingStyle,
      });
    } else if (row.operation === "消去") {
      return DeleteRecurringEventRow.parse({
        type: "deleteRecurringEvent",
        dayOfWeek: row.dayOfWeek,
      });
    } else {
      return {
        type: "no-operation",
      } satisfies NoOperationRow;
    }
  });

  return {
    after: after,
    comment: comment,
    registrationRows: sheetValues.filter(isRegistrationRow),
    modificationRows: sheetValues.filter(isModificationRow),
    deletionRows: sheetValues.filter(isDeletionRow),
  };
};

const getRecurringEventSheetRows = (sheet: GoogleAppsScript.Spreadsheet.Sheet): RecurringEventSheetRow[] => {
  const sheetRows = sheet
    .getRange("A9:G13")
    .getValues()
    .map((row) =>
      RecurringEventSheetRow.parse({
        //TODO: 2度parseしているので、1度にまとめる
        operation: row[0],
        dayOfWeek: row[1],
        startTime: row[2],
        endTime: row[3],
        restStartTime: row[4],
        restEndTime: row[5],
        workingStyle: row[6],
      }),
    );
  return sheetRows;
};

const isRegistrationRow = (
  row: RegisterRecurringEventRow | ModifyRecurringEventRow | DeleteRecurringEventRow | NoOperationRow,
): row is RegisterRecurringEventRow => row.type === "registerRecurringEvent";
const isModificationRow = (
  row: RegisterRecurringEventRow | ModifyRecurringEventRow | DeleteRecurringEventRow | NoOperationRow,
): row is ModifyRecurringEventRow => row.type === "modifyRecurringEvent";
const isDeletionRow = (
  row: RegisterRecurringEventRow | ModifyRecurringEventRow | DeleteRecurringEventRow | NoOperationRow,
): row is DeleteRecurringEventRow => row.type === "deleteRecurringEvent";
