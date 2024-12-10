import { isWeekend, set } from "date-fns";

export const isBankHoliday = (date: Date): boolean => {
  const calendarId = "c_0c278abe4ff7753bce9736b216c2d9ea4d022aa56879589f4ff71152fb7eaae8@group.calendar.google.com";
  const calendar = CalendarApp.getCalendarById(calendarId);
  const bankHolidayEvents = calendar.getEventsForDay(date);
  return bankHolidayEvents.length > 0 || isWeekend(date);
};

//NOTE: Googleスプレッドシートでは時間のみの入力がDate型として取得される際、日付部分はデフォルトで1899/12/30となるため適切な日付情報に更新する必要がある
export const mergeTimeToDate = (date: Date, time: Date): Date => {
  return set(date, { hours: time.getHours(), minutes: time.getMinutes() });
};
