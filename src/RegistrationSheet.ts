import { PartTimerProfile } from "./JobSheet";
import { createTitleFromEventInfo } from "./shift-changer";
import { EventInfo } from "./shift-changer-api";

export const insertRegistrationSheet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try {
    sheet = spreadsheet.insertSheet(`登録`, 0);
  } catch {
    throw new Error("既存の「登録」シートを使用してください");
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

export const getRegistrationInfos = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  partTimerProfile: PartTimerProfile
): EventInfo[] => {
  const registrationInfos = sheet
    .getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn())
    .getValues()
    .map((eventInfo) => {
      const date = new Date(eventInfo[0]);
      const startTimeDate= new Date(eventInfo[1]);
      const startTime = new Date(
        date.getFullYear(),
        date.getMonth(),
        date.getDate(),
        startTimeDate.getHours(),
        startTimeDate.getMinutes()
      );
      const endTimeDate = new Date(eventInfo[2]);
      const endTime = new Date(
        date.getFullYear(),
        date.getMonth(),
        date.getDate(),
        endTimeDate.getHours(),
        endTimeDate.getMinutes()
      );
      const workingStyle = eventInfo[5] as string;
      if (workingStyle === "") throw new Error("working style is not defined");
      if (eventInfo[3] === "" || eventInfo[4] === "") {
        const title = createTitleFromEventInfo({ workingStyle }, partTimerProfile);
        return { title, date, startTime, endTime };
      } else {
        const restStartTime = eventInfo[3];
        const restEndTime = eventInfo[4];
        const title = createTitleFromEventInfo({ restStartTime, restEndTime, workingStyle }, partTimerProfile);
        return { title, date, startTime, endTime };
      }
    });
  return registrationInfos;
};
