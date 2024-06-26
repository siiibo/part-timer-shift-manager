import { addMonths } from "date-fns";

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
  const twoMonthLater = addMonths(today, 2);
  const events = calendar.getEvents(today, twoMonthLater);
  events.forEach((event) => {
    const statTime = new Date(event.getStartTime().getTime());
    if (isHoliday(statTime)) {
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
