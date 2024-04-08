import { addWeeks } from "date-fns";
import { z } from "zod";

import { getConfig } from "./config";

const dayOfWeek = z
  .literal("月曜日")
  .or(z.literal("火曜日"))
  .or(z.literal("水曜日"))
  .or(z.literal("木曜日"))
  .or(z.literal("金曜日"));

export const EventInfo = z.object({
  title: z.string(),
  date: z.coerce.date(), //TODO: 日付情報だけの変数dateを消去する
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type EventInfo = z.infer<typeof EventInfo>;

const ModificationInfo = z.object({
  previousEventInfo: EventInfo,
  newEventInfo: EventInfo,
});

const RegistrationRepeatScheduleRow = z.object({
  dayOfWeek: dayOfWeek,
  startOrEndDate: z.coerce.date(),
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
type RegistrationRepeatScheduleRow = z.infer<typeof RegistrationRepeatScheduleRow>;

const getCalendar = () => {
  const { CALENDAR_ID } = getConfig();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  return calendar;
};

const isEventGuest = (event: GoogleAppsScript.Calendar.CalendarEvent, email: string) => {
  const guestEmails = event.getGuestList().map((guest) => guest.getEmail());
  return guestEmails.includes(email);
};

export const shiftChanger = (e: GoogleAppsScript.Events.DoPost) => {
  const operationType = e.parameter.operationType;
  const userEmail = e.parameter.userEmail;
  switch (operationType) {
    case "registration": {
      const registrationInfos = EventInfo.array().parse(JSON.parse(e.parameter.registrationInfos));
      registration(userEmail, registrationInfos);
      break;
    }
    case "modificationAndDeletion": {
      const modificationInfos = ModificationInfo.array().parse(JSON.parse(e.parameter.modificationInfos));
      const deletionInfos = EventInfo.array().parse(JSON.parse(e.parameter.deletionInfos));

      modification(modificationInfos, userEmail);
      deletion(deletionInfos, userEmail);
      break;
    }
    case "showEvents": {
      const startDate = new Date(e.parameter.startDate);
      const eventInfos = showEvents(userEmail, startDate);
      return JSON.stringify(eventInfos);
    }
    case "repeatSchedule": {
      const repeatScheduleRegistrationInfos = RegistrationRepeatScheduleRow.array().parse(
        JSON.parse(e.parameter.repeatScheduleModification),
      );

      repeatScheduleRegistration(repeatScheduleRegistrationInfos, userEmail);
    }
  }
  return;
};

const registration = (userEmail: string, registrationInfos: EventInfo[]) => {
  registrationInfos.forEach((registrationInfo) => {
    registerEvent(registrationInfo, userEmail);
  });
};

const registerEvent = (eventInfo: EventInfo, userEmail: string) => {
  const calendar = getCalendar();
  const [startDate, endDate] = [eventInfo.startTime, eventInfo.endTime];
  calendar.createEvent(eventInfo.title, startDate, endDate, { guests: userEmail });
};

const showEvents = (userEmail: string, startDate: Date): EventInfo[] => {
  const endDate = addWeeks(startDate, 4);
  const calendar = getCalendar();
  const events = calendar.getEvents(startDate, endDate).filter((event) => isEventGuest(event, userEmail));
  const eventInfos = events.map((event) => {
    const title = event.getTitle();
    const date = new Date(event.getStartTime().getTime());
    const startTime = new Date(event.getStartTime().getTime());
    const endTime = new Date(event.getEndTime().getTime());

    return { title, date, startTime, endTime };
  });
  return eventInfos;
};

const modification = (
  modificationInfos: {
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  }[],
  userEmail: string,
) => {
  const calendar = getCalendar();
  modificationInfos.forEach((eventInfo) => modifyEvent(eventInfo, calendar, userEmail));
};

const repeatScheduleRegistration = (
  repeatScheduleRegistrationInfos: RegistrationRepeatScheduleRow[],
  userEmail: string,
) => {
  const calendar = getCalendar();
  repeatScheduleRegistrationInfos.forEach((eventInfo) => {
    const dayOfWeek = convertJapaneseToEnglishDayOfWeek(eventInfo.dayOfWeek);
    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(dayOfWeek);
    calendar.createEventSeries(eventInfo.title, eventInfo.startTime, eventInfo.endTime, recurrence, {
      guests: userEmail,
    });
  });
};

const modifyEvent = (
  eventInfo: {
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  },
  calendar: GoogleAppsScript.Calendar.Calendar,
  userEmail: string,
) => {
  const [startDate, endDate] = [eventInfo.previousEventInfo.startTime, eventInfo.previousEventInfo.endTime];
  const newTitle = eventInfo.newEventInfo.title;
  const [newStartDate, newEndDate] = [eventInfo.newEventInfo.startTime, eventInfo.newEventInfo.endTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.setTime(newStartDate, newEndDate);
  event.setTitle(newTitle);
};

const deletion = (deletionInfos: EventInfo[], userEmail: string) => {
  const calendar = getCalendar();
  deletionInfos.forEach((eventInfo) => deleteEvent(eventInfo, calendar, userEmail));
};

const deleteEvent = (eventInfo: EventInfo, calendar: GoogleAppsScript.Calendar.Calendar, userEmail: string) => {
  const [startDate, endDate] = [eventInfo.startTime, eventInfo.endTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.deleteEvent();
};
const convertJapaneseToEnglishDayOfWeek = (dayOfWeek: string) => {
  switch (dayOfWeek) {
    case "月曜日":
      return CalendarApp.Weekday.MONDAY;
    case "火曜日":
      return CalendarApp.Weekday.TUESDAY;
    case "水曜日":
      return CalendarApp.Weekday.WEDNESDAY;
    case "木曜日":
      return CalendarApp.Weekday.THURSDAY;
    case "金曜日":
      return CalendarApp.Weekday.FRIDAY;
    default:
      throw new Error("Invalid day of the week");
  }
};
