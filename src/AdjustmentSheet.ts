import { z } from "zod";

const dateOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), z.date().optional());
const stringOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), z.string().optional());

export const DeleteAdjustmentRow = z.object({
  type: z.literal("delete"),
  endDate: z.date(),
  title: z.string(),
  startTime: z.date(),
  endTime: z.date(),
});
export type DeleteAdjustmentRow = z.infer<typeof DeleteAdjustmentRow>;

export const ModificationAdjustmentRow = z.object({
  type: z.literal("modification"),
  startDate: z.date(),
  title: z.string(),
  startTime: z.date(),
  endTime: z.date(),
  newStartTime: z.date(),
  newEndTime: z.date(),
  newRestStartTime: z.date().optional(),
  newRestEndTime: z.date().optional(),
  newWorkingStyle: z.literal("リモート").or(z.literal("出社")),
});
export type ModificationAdjustmentRow = z.infer<typeof ModificationAdjustmentRow>;

export const RegistrationAdjustmentRow = z.object({
  type: z.literal("registration"),
  startDate: z.date(),
  startTime: z.date(),
  endTime: z.date(),
  restStartTime: z.date().optional(),
  restEndTime: z.date().optional(),
  workingStyle: z.literal("リモート").or(z.literal("出社")),
});
export type RegistrationAdjustmentRow = z.infer<typeof RegistrationAdjustmentRow>;

const AdjustmentSheetRow = z.object({
  startDate: z.date(),
  title: stringOrEmptyString,
  startTime: dateOrEmptyString,
  endTime: dateOrEmptyString,
  newStartTime: dateOrEmptyString,
  newEndTime: dateOrEmptyString,
  newRestStartTime: dateOrEmptyString,
  newRestEndTime: dateOrEmptyString,
  newWorkingStyle: z.literal("リモート").or(z.literal("出社")).or(z.literal("")),
  isDelete: z.coerce.boolean(),
});
type AdjustmentSheetRow = z.infer<typeof AdjustmentSheetRow>;

const NoOperationRow = z.object({
  type: z.literal("no-operation"),
});
type NoOperationRow = z.infer<typeof NoOperationRow>;

export const insertAdjustmentSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`固定シフト`, 0);
  } catch {
    throw new Error("既存の「固定シフト」シートを使用してください");
  }
  sheet.addDeveloperMetadata(`part-timer-shift-manager-adjustment`);
  setValuesAdjustmentSheet(sheet);
};
const setValuesAdjustmentSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
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
  sheet.getRange("A9").setValue("曜日").setFontWeight("bold");
  sheet.getRange("B7").setValue(description3).setFontWeight("bold");
  sheet.getRange("B8").setValue("固定シフト変更前").setFontWeight("bold");
  sheet.getRange("E8").setValue("固定シフト変更後").setFontWeight("bold");
  sheet.getRange("J8").setValue("【削除】削除したい固定シフトを選択してください").setFontWeight("bold");
  const header2 = [
    "開始時間",
    "終了時間",
    "開始時刻",
    "終了時刻",
    "休憩開始時刻",
    "休憩終了時刻",
    "勤務形態",
    "消去対象",
  ];
  sheet.getRange(9, 3, 1, header2.length).setValues([header2]).setFontWeight("bold");

  const daysOfWeekCells = sheet.getRange("A10:A14");
  daysOfWeekCells
    .setValues([["月曜日"], ["火曜日"], ["水曜日"], ["木曜日"], ["金曜日"]])
    .setHorizontalAlignment("right");
  const workingStyleCells = sheet.getRange("I10:I14");
  const workingStyleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["リモート", "出社"], true)
    .setAllowInvalid(false)
    .setHelpText("リモート/出社 を選択してください。")
    .build();
  workingStyleCells.setDataValidation(workingStyleRule);
  const checkboxCells = sheet.getRange("J10:J14");
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .setHelpText("チェックボックス以外の入力形式は認められません。")
    .build();
  checkboxCells.setDataValidation(checkboxRule);
  sheet.setColumnWidth(1, 370);
  sheet.setColumnWidth(2, 150);
};
const getAdjustmentReSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): (DeleteAdjustmentRow | ModificationAdjustmentRow | RegistrationAdjustmentRow | NoOperationRow)[] => {
  const startDate = sheet.getRange("A5").getValue();
  const sheetValues = sheet
    .getRange("B10:J14")
    .getValues()
    .map((row) =>
      AdjustmentSheetRow.parse({
        startDate: startDate,
        title: row[0],
        startTime: row[1],
        endTime: row[2],
        newStartTime: row[3],
        newEndTime: row[4],
        newRestStartTime: row[5],
        newRestEndTime: row[6],
        newWorkingStyle: row[7],
        isDelete: row[8],
      }),
    )
    .map((row) => {
      if (row.isDelete) {
        return DeleteAdjustmentRow.parse({
          type: "delete",
          endDate: row.startDate,
          title: row.title,
          startTime: row.startTime,
          endTime: row.endTime,
        });
      } else if (!row.title && row.startDate && row.newStartTime && row.newEndTime) {
        return RegistrationAdjustmentRow.parse({
          type: "registration",
          startDate: row.startDate,
          startTime: row.newStartTime,
          endTime: row.newEndTime,
          restStartTime: row.newRestStartTime,
          restEndTime: row.newRestEndTime,
          workingStyle: row.newWorkingStyle,
        });
      } else if (row.title && row.startDate && row.newStartTime && row.newEndTime) {
        return ModificationAdjustmentRow.parse({
          type: "modification",
          startDate: row.startDate,
          title: row.title,
          startTime: row.startTime,
          endTime: row.endTime,
          newStartTime: row.newStartTime,
          newEndTime: row.newEndTime,
          newRestStartTime: row.newRestStartTime,
          newRestEndTime: row.newRestEndTime,
          newWorkingStyle: row.newWorkingStyle,
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
  row: ModificationAdjustmentRow | DeleteAdjustmentRow | RegistrationAdjustmentRow | NoOperationRow,
): row is ModificationAdjustmentRow => row.type === "modification";
const isDeletionRow = (
  row: ModificationAdjustmentRow | DeleteAdjustmentRow | RegistrationAdjustmentRow | NoOperationRow,
): row is DeleteAdjustmentRow => row.type === "delete";
const isRegistrationRow = (
  row: ModificationAdjustmentRow | DeleteAdjustmentRow | RegistrationAdjustmentRow | NoOperationRow,
): row is RegistrationAdjustmentRow => row.type === "registration";

export const getAdjustmentModificationOrDeletionOrRegistration = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): {
  registrationRows: RegistrationAdjustmentRow[];
  modificationRows: ModificationAdjustmentRow[];
  deletionRows: DeleteAdjustmentRow[];
} => {
  const sheetValues = getAdjustmentReSheetValues(sheet);
  return {
    registrationRows: sheetValues.filter(isRegistrationRow),
    modificationRows: sheetValues.filter(isModificationRow),
    deletionRows: sheetValues.filter(isDeletionRow),
  };
};
