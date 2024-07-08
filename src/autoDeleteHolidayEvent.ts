import { addWeeks } from "date-fns";

import { getConfig } from "./config";
import { isBankHoliday } from "./date-utils";

const DELETE_HOLIDAY_SHIFT_INTERVAL = 2;

export const initDeleteHolidayShift = () => {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === deleteHolidayShift.name)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger(deleteHolidayShift.name)
    .timeBased()
    .everyWeeks(DELETE_HOLIDAY_SHIFT_INTERVAL)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(0)
    .create();
};

export const deleteHolidayShift = () => {
  const { CALENDAR_ID } = getConfig();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const today = new Date();
  const twoWeekLater = addWeeks(today, DELETE_HOLIDAY_SHIFT_INTERVAL);
  const events = calendar.getEvents(today, twoWeekLater);
  events.forEach((event) => {
    const statTime = new Date(event.getStartTime().getTime());
    if (isBankHoliday(statTime)) {
      event.deleteEvent();
    }
  });
};
