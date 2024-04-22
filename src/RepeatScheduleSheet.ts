import { set } from "date-fns";
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
const WorkingStyleOrEmptyString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("リモート").or(z.literal("出勤")).optional(),
);

const OperationString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("時間変更").or(z.literal("消去")).or(z.literal("追加")).optional(),
);

const DeleteRepeatScheduleRow = z.object({
  type: z.literal("delete"),
  date: z.date(),
  oldDayOfWeek: DayOfWeekOrEmptyString,
});
type DeleteRepeatScheduleRow = z.infer<typeof DeleteRepeatScheduleRow>;

const ModificationRepeatScheduleRow = z.object({
  type: z.literal("modification"),
  after: z.date(),
  endDate: z.date(),
  oldDayOfWeek: DayOfWeekOrEmptyString,
  newDayOfWeek: DayOfWeekOrEmptyString,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出勤")),
});
type ModificationRepeatScheduleRow = z.infer<typeof ModificationRepeatScheduleRow>;

const RegistrationRepeatScheduleRow = z.object({
  type: z.literal("registration"),
  after: z.date(),
  newDayOfWeek: DayOfWeekOrEmptyString,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出勤")),
});
type RegistrationRepeatScheduleRow = z.infer<typeof RegistrationRepeatScheduleRow>;

const RepeatScheduleSheetRow = z.object({
  after: z.date(),
  operation: OperationString,
  dayOfWeek: DayOfWeekOrEmptyString,
  startTime: DateOrEmptyString,
  endTime: DateOrEmptyString,
  restStartTime: DateOrEmptyString,
  restEndTime: DateOrEmptyString,
  workingStyle: WorkingStyleOrEmptyString,
});
type RepeatScheduleSheetRow = z.infer<typeof RepeatScheduleSheetRow>;

const NoOperationRow = z.object({
  type: z.literal("no-operation"),
});
type NoOperationRow = z.infer<typeof NoOperationRow>;

export const insertRepeatScheduleSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`固定シフト`, 0);
  } catch {
    throw new Error("既存の「固定シフト」シートを使用してください");
  }
  sheet.addDeveloperMetadata(`part-timer-shift-manager-repeatSchedule`);
  setValuesRepeatScheduleSheet(sheet);
};
const setValuesRepeatScheduleSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
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
const getRepeatScheduleReSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): (DeleteRepeatScheduleRow | ModificationRepeatScheduleRow | RegistrationRepeatScheduleRow | NoOperationRow)[] => {
  const after = sheet.getRange("A5").getValue();
  const sheetValues = sheet
    .getRange("A9:G13")
    .getValues()
    .map((row) =>
      RepeatScheduleSheetRow.parse({
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
        return DeleteRepeatScheduleRow.parse({
          type: "delete",
          after: row.after,
          dayOfWeek: row.dayOfWeek,
        });
      } else if (row.operation === "追加" && row.dayOfWeek && row.startTime && row.endTime) {
        const nextDate = getNextDayOfWeek(row.after, row.dayOfWeek);
        const startTime = mergeTimeToDate(nextDate, row.startTime);
        const endTime = mergeTimeToDate(nextDate, row.endTime);
        return RegistrationRepeatScheduleRow.parse({
          type: "registration",
          after: row.after,
          dayOfWeek: row.dayOfWeek,
          startTime: startTime,
          endTime: endTime,
          restStartTime: row.restStartTime,
          restEndTime: row.restEndTime,
          workingStyle: row.workingStyle,
        });
      } else if (row.operation === "時間変更" && row.dayOfWeek && row.startTime && row.endTime && row.after) {
        const after = getNextDayOfWeek(row.after, row.dayOfWeek);
        const startTime = mergeTimeToDate(after, row.startTime);
        const endTime = mergeTimeToDate(after, row.endTime);
        return ModificationRepeatScheduleRow.parse({
          type: "modification",
          after: row.after,
          dayOfWeek: row.dayOfWeek,
          startTime: startTime,
          endTime: endTime,
          restStartTime: row.restStartTime,
          restEndTime: row.restEndTime,
          workingStyle: row.workingStyle,
        });
      } else {
        return NoOperationRow.parse({
          type: "no-operation",
        });
      }
    });
  return sheetValues;
};

const isModificationRow = (
  row: ModificationRepeatScheduleRow | DeleteRepeatScheduleRow | RegistrationRepeatScheduleRow | NoOperationRow,
): row is ModificationRepeatScheduleRow => row.type === "modification";
const isDeletionRow = (
  row: ModificationRepeatScheduleRow | DeleteRepeatScheduleRow | RegistrationRepeatScheduleRow | NoOperationRow,
): row is DeleteRepeatScheduleRow => row.type === "delete";
const isRegistrationRow = (
  row: ModificationRepeatScheduleRow | DeleteRepeatScheduleRow | RegistrationRepeatScheduleRow | NoOperationRow,
): row is RegistrationRepeatScheduleRow => row.type === "registration";

export const getRepeatScheduleModificationOrDeletionOrRegistration = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): {
  registrationRows: RegistrationRepeatScheduleRow[];
  modificationRows: ModificationRepeatScheduleRow[];
  deletionRows: DeleteRepeatScheduleRow[];
} => {
  const sheetValues = getRepeatScheduleReSheetValues(sheet);
  return {
    registrationRows: sheetValues.filter(isRegistrationRow),
    modificationRows: sheetValues.filter(isModificationRow),
    deletionRows: sheetValues.filter(isDeletionRow),
  };
};
//NOTE: Googleスプレッドシートでは時間のみの入力がDate型として取得される際、日付部分はデフォルトで1899/12/30となるため適切な日付情報に更新する必要がある
const mergeTimeToDate = (date: Date, time: Date): Date => {
  return set(date, { hours: time.getHours(), minutes: time.getMinutes() });
};
//NOTE: 仕様的にstartTimeの日付に最初の予定が指定されるため、指定された日付の後で一番近い指定曜日の日付に変更する
const getNextDayOfWeek = (startDate: Date, newDayOfWeek: string): Date => {
  const daysOfWeek = ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"];
  const targetDayOfWeek = daysOfWeek.indexOf(newDayOfWeek.toLowerCase());
  if (targetDayOfWeek === -1) {
    throw new Error("Invalid day of week specified");
  }

  const currentDayOfWeek = startDate.getDay();
  const daysToAdd = (targetDayOfWeek + 7 - currentDayOfWeek) % 7;
  const nextDate = new Date(startDate);
  nextDate.setDate(startDate.getDate() + daysToAdd);
  return nextDate;
};
