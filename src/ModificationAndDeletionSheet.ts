import { z } from "zod";

import { Comment, DateAfterNow, DateOrEmptyString } from "./common.schema";
import { mergeTimeToDate } from "./date-utils";

const ModificationRow = z.object({
  type: z.literal("modification"),
  title: z.string(),
  startTime: DateAfterNow,
  endTime: DateAfterNow,
  newStartTime: DateAfterNow,
  newEndTime: DateAfterNow,
  newRestStartTime: z.date().optional(),
  newRestEndTime: z.date().optional(),
  newWorkingStyle: z.literal("出社").or(z.literal("リモート")),
});
type ModificationRow = z.infer<typeof ModificationRow>;

const DeletionRow = z.object({
  type: z.literal("deletion"),
  title: z.string(),
  startTime: DateAfterNow,
  endTime: DateAfterNow,
});
type DeletionRow = z.infer<typeof DeletionRow>;

const ModificationOrDeletionSheetRow = z
  .object({
    title: z.string(),
    date: z.date(),
    startTime: z.date(),
    endTime: z.date(),
    newDate: DateOrEmptyString,
    newStartTime: DateOrEmptyString,
    newEndTime: DateOrEmptyString,
    newRestStartTime: DateOrEmptyString,
    newRestEndTime: DateOrEmptyString,
    newWorkingStyle: z.literal("出社").or(z.literal("リモート")).or(z.literal("")),
    isDeletionTarget: z.coerce.boolean(),
  })
  .refine(
    (data) => {
      if (data.newRestStartTime && data.newRestEndTime) {
        return data.newRestStartTime < data.newRestEndTime;
      }
      return true;
    },
    {
      message: "休憩時間の開始時間が終了時間よりも前になるようにしてください",
    },
  );

const NoOperationRow = z.object({
  type: z.literal("no-operation"),
});
type NoOperationRow = z.infer<typeof NoOperationRow>;

type ModificationOrDeletionSheetValues = {
  comment: Comment;
  modificationRows: ModificationRow[];
  deletionRows: DeletionRow[];
};

export const insertModificationAndDeletionSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet: GoogleAppsScript.Spreadsheet.Sheet;
  try {
    sheet = spreadsheet.insertSheet("単発シフト変更・削除", 0);
  } catch {
    throw new Error("既存の「単発シフト変更・削除」シートを使用してください");
  }
  sheet.addDeveloperMetadata("part-timer-shift-manager-modificationAndDeletion");
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

export const getModificationOrDeletionSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): ModificationOrDeletionSheetValues => {
  const { modificationRows, deletionRows } = getModificationOrDeletionRows(sheet);
  const comment = Comment.parse(sheet.getRange("A2").getValue());
  return {
    comment,
    modificationRows,
    deletionRows,
  };
};

const getModificationOrDeletionRows = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): { modificationRows: ModificationRow[]; deletionRows: DeletionRow[] } => {
  const sheetValues = sheet
    .getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn())
    .getValues()
    .map((row) =>
      ModificationOrDeletionSheetRow.parse({
        //TODO: 2度parseしているので、1度にまとめる
        title: row[0],
        date: row[1],
        startTime: row[2],
        endTime: row[3],
        newDate: row[4], // 未入力の場合は空文字、それ以外の場合はDate型が渡ってくる
        newStartTime: row[5], // 未入力の場合は空文字、それ以外の場合はDate型が渡ってくる
        newEndTime: row[6], // 未入力の場合は空文字、それ以外の場合はDate型が渡ってくる
        newRestStartTime: row[7],
        newRestEndTime: row[8],
        newWorkingStyle: row[9],
        isDeletionTarget: row[10],
      }),
    )
    .map((row) => {
      if (row.isDeletionTarget) {
        const startTime = mergeTimeToDate(row.date, row.startTime);
        const endTime = mergeTimeToDate(row.date, row.endTime);
        return DeletionRow.parse({
          type: "deletion",
          title: row.title,
          startTime: startTime,
          endTime: endTime,
        });
      }
      if (row.newDate && row.newStartTime && row.newEndTime) {
        const startTime = mergeTimeToDate(row.date, row.startTime);
        const endTime = mergeTimeToDate(row.date, row.endTime);
        const newStartTime = mergeTimeToDate(row.newDate, row.newStartTime);
        const newEndTime = mergeTimeToDate(row.newDate, row.newEndTime);
        return ModificationRow.parse({
          type: "modification",
          title: row.title,
          startTime: startTime,
          endTime: endTime,
          newStartTime: newStartTime,
          newEndTime: newEndTime,
          newRestStartTime: row.newRestEndTime,
          newRestEndTime: row.newRestEndTime,
          newWorkingStyle: row.newWorkingStyle,
        });
      }
      return NoOperationRow.parse({
        type: "no-operation",
      });
    });
  return {
    modificationRows: sheetValues.filter(isModificationRow),
    deletionRows: sheetValues.filter(isDeletionRow),
  };
};

const isModificationRow = (row: ModificationRow | DeletionRow | NoOperationRow): row is ModificationRow =>
  row.type === "modification";
const isDeletionRow = (row: ModificationRow | DeletionRow | NoOperationRow): row is DeletionRow =>
  row.type === "deletion";
