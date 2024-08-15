import { getDate, getMonth, isWeekend, set } from "date-fns";

const isHoliday = (day: Date): boolean => {
  const calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
  const calendar = CalendarApp.getCalendarById(calendarId);
  const holidayEvents = calendar.getEventsForDay(day);
  return holidayEvents.length > 0;
};

export const isBankHoliday = (date: Date): boolean => {
  // 12/31
  if (getMonth(date) === 11 && getDate(date) === 31) {
    return true;
  }
  // 1/1 ~ 1/3
  if (getMonth(date) === 0) {
    if (getDate(date) === 1 || getDate(date) === 2 || getDate(date) === 3) {
      return true;
    }
  }
  if (isHoliday(date) || isWeekend(date)) {
    return true;
  }
  return false;
};

//NOTE: Googleスプレッドシートでは時間のみの入力がDate型として取得される際、日付部分はデフォルトで1899/12/30となるため適切な日付情報に更新する必要がある
export const mergeTimeToDate = (date: Date, time: Date): Date => {
  return set(date, { hours: time.getHours(), minutes: time.getMinutes() });
};
