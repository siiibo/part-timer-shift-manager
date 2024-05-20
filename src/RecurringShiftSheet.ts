import { z } from "zod";

const DateOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), z.date().optional());
const DayOfWeekOrEmptyString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z
    .literal("月曜日")
    .or(z.literal("火曜日"))
    .or(z.literal("水曜日"))
    .or(z.literal("木曜日"))
    .or(z.literal("金曜日"))
    .optional(),
);
const DayOfWeek = z
  .literal("月曜日")
  .or(z.literal("火曜日"))
  .or(z.literal("水曜日"))
  .or(z.literal("木曜日"))
  .or(z.literal("金曜日"));
const WorkingStyleOrEmptyString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("リモート").or(z.literal("出勤")).optional(),
);

const OperationString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("時間変更").or(z.literal("消去")).or(z.literal("追加")).optional(),
);

const DeletionRecurringEventRow = z.object({
  type: z.literal("deletion"),
  after: z.date(),
  dayOfWeek: DayOfWeek,
});
type DeletionRecurringEventRow = z.infer<typeof DeletionRecurringEventRow>;

const ModificationRecurringEventRow = z.object({
  type: z.literal("modification"),
  after: z.date(),
  dayOfWeek: DayOfWeek,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出勤")),
});
type ModificationRecurringEventRow = z.infer<typeof ModificationRecurringEventRow>;

const RegistrationRecurringEventRow = z.object({
  type: z.literal("registration"),
  after: z.date(),
  dayOfWeek: DayOfWeek,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出勤")),
});
type RegistrationRecurringEventRow = z.infer<typeof RegistrationRecurringEventRow>;

const RecurringEventSheetRow = z.object({
  after: z.date(),
  operation: OperationString,
  dayOfWeek: DayOfWeekOrEmptyString,
  startTime: DateOrEmptyString,
  endTime: DateOrEmptyString,
  restStartTime: DateOrEmptyString,
  restEndTime: DateOrEmptyString,
  workingStyle: WorkingStyleOrEmptyString,
});
type RecurringEventSheetRow = z.infer<typeof RecurringEventSheetRow>;

type NoOperationRow = {
  type: "no-operation";
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

const setValuesRecurringEventSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
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
    .requireValueInList(["時間変更", "消去", "追加"], true)
    .setAllowInvalid(false)
    .setHelpText("時間変更, 削除, 登録 を選択してください。")
    .build();
  operationCells.setDataValidation(operationRule);
  sheet.getRange("B9:B13").setValues([["月曜日"], ["火曜日"], ["水曜日"], ["木曜日"], ["金曜日"]]);

  const workingStyleCells = sheet.getRange("G9:G13");
  const workingStyleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リモート", "出勤"], true)
    .setAllowInvalid(false)
    .setHelpText("リモート/出社 を選択してください。")
    .build();
  workingStyleCells.setDataValidation(workingStyleRule);

  sheet.setColumnWidth(1, 370);
  sheet.setColumnWidth(2, 150);
};

const getRecurringEventSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): (DeletionRecurringEventRow | ModificationRecurringEventRow | RegistrationRecurringEventRow | NoOperationRow)[] => {
  const after = sheet.getRange("A5").getValue();
  const sheetValues = sheet
    .getRange("A9:G13")
    .getValues()
    .map((row) =>
      RecurringEventSheetRow.parse({
        after: after,
        operation: row[0],
        dayOfWeek: row[1],
        startTime: row[2],
        endTime: row[3],
        restStartTime: row[4],
        restEndTime: row[5],
        workingStyle: row[6],
      }),
    )
    .map((row) => {
      if (row.operation === "消去") {
        return DeletionRecurringEventRow.parse({
          type: "deletion",
          after: row.after,
          dayOfWeek: row.dayOfWeek,
        });
      } else if (row.operation === "追加" && row.dayOfWeek && row.startTime && row.endTime) {
        return RegistrationRecurringEventRow.parse({
          type: "registration",
          after: row.after,
          dayOfWeek: row.dayOfWeek,
          startTime: row.startTime,
          endTime: row.endTime,
          restStartTime: row.restStartTime,
          restEndTime: row.restEndTime,
          workingStyle: row.workingStyle,
        });
      } else if (row.operation === "時間変更" && row.dayOfWeek && row.startTime && row.endTime) {
        return ModificationRecurringEventRow.parse({
          type: "modification",
          after: row.after,
          dayOfWeek: row.dayOfWeek,
          startTime: row.startTime,
          endTime: row.endTime,
          restStartTime: row.restStartTime,
          restEndTime: row.restEndTime,
          workingStyle: row.workingStyle,
        });
      } else {
        return {
          type: "no-operation",
        } satisfies NoOperationRow;
      }
    });
  return sheetValues;
};

const isModificationRow = (
  row: ModificationRecurringEventRow | DeletionRecurringEventRow | RegistrationRecurringEventRow | NoOperationRow,
): row is ModificationRecurringEventRow => row.type === "modification";
const isDeletionRow = (
  row: ModificationRecurringEventRow | DeletionRecurringEventRow | RegistrationRecurringEventRow | NoOperationRow,
): row is DeletionRecurringEventRow => row.type === "deletion";
const isRegistrationRow = (
  row: ModificationRecurringEventRow | DeletionRecurringEventRow | RegistrationRecurringEventRow | NoOperationRow,
): row is RegistrationRecurringEventRow => row.type === "registration";

export const getRecurringEventModificationOrDeletionOrRegistration = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): {
  registrationRows: RegistrationRecurringEventRow[];
  modificationRows: ModificationRecurringEventRow[];
  deletionRows: DeletionRecurringEventRow[];
} => {
  const sheetValues = getRecurringEventSheetValues(sheet);
  return {
    registrationRows: sheetValues.filter(isRegistrationRow),
    modificationRows: sheetValues.filter(isModificationRow),
    deletionRows: sheetValues.filter(isDeletionRow),
  };
};
