import { set } from "date-fns";
import { z } from "zod";

const dateOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), z.date().optional());
const dayOfWeekOrEmptyString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z
    .literal("月曜日")
    .or(z.literal("火曜日"))
    .or(z.literal("水曜日"))
    .or(z.literal("木曜日"))
    .or(z.literal("金曜日"))
    .optional(),
);
const workingStyleOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), z.string().optional());

const DeleteRepeatScheduleRow = z.object({
  type: z.literal("delete"),
  startOrEndDate: z.date(),
  oldDayOfWeek: dayOfWeekOrEmptyString,
});
type DeleteRepeatScheduleRow = z.infer<typeof DeleteRepeatScheduleRow>;

const ModificationRepeatScheduleRow = z.object({
  type: z.literal("modification"),
  startDate: z.date(),
  oldDayOfWeek: dayOfWeekOrEmptyString,
  newDayOfWeek: dayOfWeekOrEmptyString,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出勤")),
});
type ModificationRepeatScheduleRow = z.infer<typeof ModificationRepeatScheduleRow>;

const RegistrationRepeatScheduleRow = z.object({
  type: z.literal("registration"),
  startDate: z.date(),
  newDayOfWeek: dayOfWeekOrEmptyString,
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出勤")),
});
type RegistrationRepeatScheduleRow = z.infer<typeof RegistrationRepeatScheduleRow>;

const RepeatScheduleSheetRow = z.object({
  startDate: z.date(),
  oldDayOfWeek: dayOfWeekOrEmptyString,
  newDayOfWeek: dayOfWeekOrEmptyString,
  startTime: dateOrEmptyString,
  endTime: dateOrEmptyString,
  restStartTime: dateOrEmptyString,
  restEndTime: dateOrEmptyString,
  workingStyle: workingStyleOrEmptyString,
  isDelete: z.coerce.boolean(),
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
  sheet.getRange("B8").setValue("固定シフト変更後").setFontWeight("bold");
  sheet.getRange("H8").setValue("【削除】削除したい固定シフトを選択してください").setFontWeight("bold");
  const header2 = [
    "追加・変更・消去する曜日を選択",
    "変更後の曜日",
    "開始時刻",
    "終了時刻",
    "休憩開始時刻",
    "休憩終了時刻",
    "勤務形態",
    "消去対象",
  ];
  sheet.getRange(9, 1, 1, header2.length).setValues([header2]).setFontWeight("bold");

  const dayOfWeekCells = sheet.getRange("A10:B14");
  const dayOfWeekRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["月曜日", "火曜日", "水曜日", "木曜日", "金曜日"])
    .setAllowInvalid(false)
    .setHelpText("曜日を選択さしてください")
    .build();
  dayOfWeekCells.setDataValidation(dayOfWeekRule);

  const workingStyleCells = sheet.getRange("G10:G14");
  const workingStyleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リモート", "出勤"], true)
    .setAllowInvalid(false)
    .setHelpText("リモート/出社 を選択してください。")
    .build();
  workingStyleCells.setDataValidation(workingStyleRule);

  workingStyleCells.setDataValidation(workingStyleRule);
  const checkboxCells = sheet.getRange("H10:H14");
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .setHelpText("チェックボックス以外の入力形式は認められません。")
    .build();
  checkboxCells.setDataValidation(checkboxRule);
  sheet.setColumnWidth(1, 370);
  sheet.setColumnWidth(2, 150);
};
const getRepeatScheduleReSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): (DeleteRepeatScheduleRow | ModificationRepeatScheduleRow | RegistrationRepeatScheduleRow | NoOperationRow)[] => {
  const startDate = sheet.getRange("A5").getValue();
  const sheetValues = sheet
    .getRange("A10:H14")
    .getValues()
    .map((row) =>
      RepeatScheduleSheetRow.parse({
        startDate: startDate,
        oldDayOfWeek: row[0],
        newDayOfWeek: row[1],
        startTime: row[2],
        endTime: row[3],
        restStartTime: row[4],
        restEndTime: row[5],
        workingStyle: row[6],
        isDelete: row[7],
      }),
    )
    .map((row) => {
      console.log(row);
      if (row.isDelete) {
        return DeleteRepeatScheduleRow.parse({
          type: "delete",
          startOrEndDate: row.startDate,
          oldDayOfWeek: row.oldDayOfWeek,
        });
      } else if (!row.oldDayOfWeek && row.startTime && row.endTime) {
        const startTime = mergeTimeToDate(row.startDate, row.startTime);
        const endTime = mergeTimeToDate(row.startDate, row.endTime);
        return RegistrationRepeatScheduleRow.parse({
          type: "registration",
          startDate: row.startDate,
          newDayOfWeek: row.newDayOfWeek,
          startTime: startTime,
          endTime: endTime,
          restStartTime: row.restStartTime,
          restEndTime: row.restEndTime,
          workingStyle: row.workingStyle,
        });
      } else if (row.oldDayOfWeek && row.startTime && row.endTime) {
        const startTime = mergeTimeToDate(row.startDate, row.startTime);
        const endTime = mergeTimeToDate(row.startDate, row.endTime);
        return ModificationRepeatScheduleRow.parse({
          type: "modification",
          startDate: row.startDate,
          oldDayOfWeek: row.oldDayOfWeek,
          newDayOfWeek: row.newDayOfWeek,
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
  console.log(sheetValues);
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
