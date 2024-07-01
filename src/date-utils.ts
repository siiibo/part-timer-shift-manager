import { getDate, getMonth } from "date-fns";

const isHoliday = (day: Date): boolean => {
  const calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
  const calendar = CalendarApp.getCalendarById(calendarId);
  const holidayEvents = calendar.getEventsForDay(day);
  return holidayEvents.length > 0;
};

export const isHolidayOrSpecialDate = (date: Date): boolean => {
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
  if (isHoliday(date)) {
    return true;
  }
  return false;
};
