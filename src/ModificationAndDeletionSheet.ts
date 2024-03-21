import { set } from "date-fns";
import { z } from "zod";

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
export const Modification = z.object({
  type: z.literal("modification"),
  title: z.string(),
  date: z.coerce.date(),
  startTime: z.coerce.date().min(new Date(), { message: "過去の時間にシフト変更はできません" }),
  endTime: z.coerce.date(),
  newDate: z.coerce.date(),
  newStartTime: z.coerce.date().min(new Date(), { message: "過去の時間にシフト変更はできません" }),
  newEndTime: z.coerce.date(),
  newRestStartTime: z.coerce.date().optional(),
  newRestEndTime: z.coerce.date().optional(),
  newWorkingStyle: z.literal("出社").or(z.literal("リモート")),
});
export type Modification = z.infer<typeof Modification>;
const Deletion = z.object({
  type: z.literal("deletion"),
  title: z.string(),
  date: z.coerce.date().min(new Date(), { message: "過去の時間はシフト削除はできません" }),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
type Deletion = z.infer<typeof Deletion>;
export type ModificationAndDeletionSheetRow = Modification | Deletion;

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
const getModificationAndDeletionSheetValues = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): ModificationAndDeletionSheetRow[] => {
  const sheetValues = sheet
    .getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn())
    .getValues()
    .map((row) => {
      const deletionFlag = row[10];
      if (deletionFlag) {
        const date = row[1];
        const startTime = set(date, {
          hours: row[2].getHours(),
          minutes: row[2].getMinutes(),
        });
        const endTime = set(date, { hours: row[3].getHours(), minutes: row[3].getMinutes() });
        return Deletion.parse({
          type: "deletion",
          title: row[0],
          date: row[1],
          startTime: startTime,
          endTime: endTime,
        });
      } else {
        const date = row[1];
        const startTime = set(date, {
          hours: row[2].getHours(),
          minutes: row[2].getMinutes(),
        });
        const endTime = set(date, { hours: row[3].getHours(), minutes: row[3].getMinutes() });
        const newDate = row[4];
        const newStartTime = set(newDate, {
          hours: row[5].getHours(),
          minutes: row[5].getMinutes(),
        });
        const newEndTime = set(newDate, {
          hours: row[6].getHours(),
          minutes: row[6].getMinutes(),
        });
        return Modification.parse({
          type: "modification",
          title: row[0],
          date: row[1],
          startTime: startTime,
          endTime: endTime,
          newDate: newDate,
          newStartTime: newStartTime,
          newEndTime: newEndTime,
          newRestStartTime: row[7] === "" ? undefined : row[7],
          newRestEndTime: row[8] === "" ? undefined : row[8],
          newWorkingStyle: row[9],
        });
      }
    });
  return sheetValues;
};
const isModification = (row: ModificationAndDeletionSheetRow): row is Modification => row.type === "modification";
const isDeletion = (row: ModificationAndDeletionSheetRow): row is Deletion => row.type === "deletion";

const getModification = (sheetValues: ModificationAndDeletionSheetRow[]): Modification[] => {
  return sheetValues.filter(isModification);
};
const getDeletion = (sheetValues: ModificationAndDeletionSheetRow[]): Deletion[] => {
  return sheetValues.filter(isDeletion);
};

export const getModificationOrDeletion = (sheet: GoogleAppsScript.Spreadsheet.Sheet): [Modification[], Deletion[]] => {
  const sheetValues = getModificationAndDeletionSheetValues(sheet);
  return [getModification(sheetValues), getDeletion(sheetValues)];
};
