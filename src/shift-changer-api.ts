import { addWeeks, endOfDay, format, nextDay, previousDay, set, startOfDay, subHours } from "date-fns";
import { z } from "zod";

import { getConfig } from "./config";

const DayOfWeek = z
  .literal("月曜日")
  .or(z.literal("火曜日"))
  .or(z.literal("水曜日"))
  .or(z.literal("木曜日"))
  .or(z.literal("金曜日"));
type DayOfWeek = z.infer<typeof DayOfWeek>;

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

const RegisterRecurringEventRequest = z.object({
  after: z.coerce.date(),
  events: z
    .object({
      dayOfWeek: DayOfWeek,
      title: z.string(),
      startTime: z.coerce.date(),
      endTime: z.coerce.date(),
    })
    .array(),
});
type RegisterRecurringEventRequest = z.infer<typeof RegisterRecurringEventRequest>;

const DeletionRecurringEvent = z.object({
  after: z.coerce.date(),
  dayOfWeeks: DayOfWeek.array(),
});
type DeletionRecurringEvent = z.infer<typeof DeletionRecurringEvent>;

type DeleteRecurringEventResponse = {
  responseCode: number;
  comment: string;
};

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
    case "registerRecurringEvent": {
      const registerRecurringEventRequest = RegisterRecurringEventRequest.parse(
        JSON.parse(e.parameter.recurringEventRegistration),
      );
      registerRecurringEvent(registerRecurringEventRequest, userEmail);
      break;
    }
    case "deleteRecurringEvent": {
      const deletionRecurringEvents = DeletionRecurringEvent.parse(JSON.parse(e.parameter.recurringEventDeletion));

      return JSON.stringify(deleteRecurringEvent(deletionRecurringEvents, userEmail));
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

const registerRecurringEvent = ({ after, events }: RegisterRecurringEventRequest, userEmail: string) => {
  const calendar = getCalendar();
  events.forEach(({ title, startTime, endTime, dayOfWeek }) => {
    const recurrenceStartDate = getNextDay(after, dayOfWeek);
    const eventStartTime = mergeTimeToDate(recurrenceStartDate, startTime);
    const eventEndTime = mergeTimeToDate(recurrenceStartDate, endTime);
    const englishDayOfWeek = convertJapaneseToEnglishDayOfWeek(dayOfWeek);

    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(englishDayOfWeek);
    calendar.createEventSeries(title, eventStartTime, eventEndTime, recurrence, {
      guests: userEmail,
    });
  });
};

const deleteRecurringEvent = (
  { dayOfWeeks, after }: DeletionRecurringEvent,
  userEmail: string,
): DeleteRecurringEventResponse => {
  const calendarId = getConfig().CALENDAR_ID;
  const advancedCalendar = Calendar.Events;
  if (advancedCalendar === undefined) return { responseCode: 400, comment: "カレンダーの取得に失敗しました" };

  const eventItems = dayOfWeeks
    .map((dayOfWeek) => {
      //NOTE: 仕様的にstartTimeの日付に最初の予定が指定されるため、指定された日付の後で一番近い指定曜日の日付に変更する
      const untilDate = getPreviousDay(after, dayOfWeek);
      const events =
        advancedCalendar.list(calendarId, {
          timeMin: startOfDay(untilDate).toISOString(),
          timeMax: endOfDay(untilDate).toISOString(),
          singleEvents: true,
          orderBy: "startTime",
          maxResults: 1,
          q: userEmail,
        }).items ?? [];
      const recurringEventId = events[0]?.recurringEventId;
      return recurringEventId ? { recurringEventId, untilDate } : undefined;
    })
    .filter(isNotUndefined);
  if (eventItems.length === 0) return { responseCode: 400, comment: "消去するイベントの取得に失敗しました" };

  const detailedEventItems = eventItems.map(({ recurringEventId, untilDate }) => {
    const eventDetail = advancedCalendar.get(calendarId, recurringEventId);
    return { eventDetail, untilDate, recurringEventId };
  });

  detailedEventItems.forEach(({ eventDetail, untilDate, recurringEventId }) => {
    if (!eventDetail.start?.dateTime || !eventDetail.end?.dateTime) return;
    const untilTime = mergeTimeToDate(untilDate, new Date(eventDetail.start.dateTime));
    const untilTimeUTC = convertJSTToUTC(untilTime);

    const data = {
      summary: eventDetail.summary,
      attendees: [{ email: userEmail }],
      start: {
        dateTime: eventDetail.start.dateTime,
        timeZone: "Asia/Tokyo",
      },
      end: {
        dateTime: eventDetail.end.dateTime,
        timeZone: "Asia/Tokyo",
      },
      recurrence: ["RRULE:FREQ=WEEKLY;UNTIL=" + format(untilTimeUTC, "yyyyMMdd'T'HHmmss'Z'")],
    };
    advancedCalendar.update(data, calendarId, recurringEventId);
  });

  return { responseCode: 200, comment: "イベントの消去が成功しました" };
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

const convertJapaneseToEnglishDayOfWeek = (dayOfWeek: DayOfWeek) => {
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

const convertJapaneseToNumberDayOfWeek = (dayOfWeek: DayOfWeek) => {
  switch (dayOfWeek) {
    case "月曜日":
      return 1;
    case "火曜日":
      return 2;
    case "水曜日":
      return 3;
    case "木曜日":
      return 4;
    case "金曜日":
      return 5;
    default:
      throw new Error("Invalid day of the week");
  }
};

const getNextDay = (date: Date, dayOfWeek: DayOfWeek): Date => {
  const targetDayOfWeek = convertJapaneseToNumberDayOfWeek(dayOfWeek);
  const nextDate = nextDay(date, targetDayOfWeek);

  return nextDate;
};

const getPreviousDay = (date: Date, dayOfWeek: DayOfWeek): Date => {
  const targetDayOfWeek = convertJapaneseToNumberDayOfWeek(dayOfWeek);
  const previousDate = previousDay(date, targetDayOfWeek);

  return previousDate;
};

const isNotUndefined = <T>(value: T | undefined): value is T => {
  return value !== undefined;
};

const mergeTimeToDate = (date: Date, time: Date): Date => {
  return set(date, { hours: time.getHours(), minutes: time.getMinutes() });
};

const convertJSTToUTC = (date: Date): string => {
  const UTCTime = subHours(date, 9);
  return UTCTime.toISOString();
};
