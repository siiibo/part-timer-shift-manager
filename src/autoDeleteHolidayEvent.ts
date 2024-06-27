import { addMonths, getDate, getMonth } from "date-fns";

import { getConfig } from "./config";

export const initDeletionHolidayShift = () => {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === deleteHolidayShift.name)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger(deleteHolidayShift.name)
    .timeBased()
    .everyWeeks(2)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(0)
    .create();
};

export const deleteHolidayShift = () => {
  const { CALENDAR_ID } = getConfig();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const today = new Date();
  const twoMonthLater = addMonths(today, 10);
  const events = calendar.getEvents(today, twoMonthLater);
  events.forEach((event) => {
    const statTime = new Date(event.getStartTime().getTime());
    if (isHolidayOrSpecialDate(statTime)) {
      event.deleteEvent();
    }
  });
};

const isHoliday = (day: Date): boolean => {
  const calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
  const calendar = CalendarApp.getCalendarById(calendarId);
  const holidayEvents = calendar.getEventsForDay(day);
  return holidayEvents.length > 0;
};

const isHolidayOrSpecialDate = (date: Date): boolean => {
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
