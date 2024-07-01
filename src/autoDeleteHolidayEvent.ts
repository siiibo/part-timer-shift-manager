import { addWeeks } from "date-fns";

import { getConfig } from "./config";
import { isHolidayOrSpecialDate } from "./date-utils";

export const initDeleteHolidayShift = () => {
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
  const twoWeekLater = addWeeks(today, 2);
  const events = calendar.getEvents(today, twoWeekLater);
  events.forEach((event) => {
    const statTime = new Date(event.getStartTime().getTime());
    if (isHolidayOrSpecialDate(statTime)) {
      event.deleteEvent();
    }
  });
};
