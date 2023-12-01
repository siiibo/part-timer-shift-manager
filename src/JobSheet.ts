import { getConfig } from "./config";

export type PartTimerProfile = {
  job: string;
  lastName: string;
  email: string;
  managerEmails: string[];
};
export const getPartTimerProfile = (userEmail: string): PartTimerProfile => {
  const { JOB_SHEET_URL } = getConfig();
  const sheet = SpreadsheetApp.openByUrl(JOB_SHEET_URL).getSheetByName("シート1");
  if (!sheet) throw new Error("SHEET is not defined");
  const partTimerProfiles = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues()
    .map((row) => ({
      job: row[0] as string,
      // \u3000は全角空白
      lastName: row[1].split(/(\s|\u3000)+/)[0] as string,
      email: row[2] as string,
      managerEmails: row[3] === "" ? [] : (row[3] as string).replaceAll(/\s/g, "").split(","),
    }));

  const partTimerProfile = partTimerProfiles.find(({ email }) => {
    return email === userEmail;
  });
  if (partTimerProfile === undefined) throw new Error("no part timer information for the email");

  return partTimerProfile;
};
