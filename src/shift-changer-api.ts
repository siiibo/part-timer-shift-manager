import { addWeeks, endOfDay, format, nextDay, previousDay, set, startOfDay, subHours } from "date-fns";
import { z } from "zod";
import { zu } from "zod_utilz";

import { getConfig } from "./config";

const DayOfWeek = z
  .literal("月曜日")
  .or(z.literal("火曜日"))
  .or(z.literal("水曜日"))
  .or(z.literal("木曜日"))
  .or(z.literal("金曜日"));
type DayOfWeek = z.infer<typeof DayOfWeek>;

export const Event = z.object({
  title: z.string(),
  date: z.coerce.date(), //TODO: 日付情報だけの変数dateを消去する
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type Event = z.infer<typeof Event>;

const RegisterEventRequest = z.object({
  operationType: z.literal("registerEvent"),
  userEmail: z.string(),
  registrationEvents: zu.stringToJSON().pipe(Event.array()),
});

const ModifyAndDeleteEventRequest = z.object({
  operationType: z.literal("modifyAndDeleteEvent"),
  userEmail: z.string(),
  modificationEvents: zu.stringToJSON().pipe(
    z
      .object({
        previousEvent: Event,
        newEvent: Event,
      })
      .array(),
  ),
  deletionEvents: zu.stringToJSON().pipe(Event.array()),
});
type ModifyAndDeleteEventRequest = z.infer<typeof ModifyAndDeleteEventRequest>;

const ShowEventRequest = z.object({
  operationType: z.literal("showEvents"),
  userEmail: z.string(),
  startDate: z.coerce.date(),
});
type ShowEventRequest = z.infer<typeof ShowEventRequest>;

const RegisterRecurringEventRequest = z.object({
  operationType: z.literal("registerRecurringEvent"),
  userEmail: z.string(),
  registrationRecurringEvents: zu.stringToJSON().pipe(
    z.object({
      after: z.coerce.date(),
      events: z
        .object({
          dayOfWeek: DayOfWeek,
          title: z.string(),
          startTime: z.coerce.date(),
          endTime: z.coerce.date(),
        })
        .array(),
    }),
  ),
});
export type RegisterRecurringEventRequest = z.infer<typeof RegisterRecurringEventRequest>;

const ModifyRecurringEventRequest = z.object({
  operationType: z.literal("modifyRecurringEvent"),
  userEmail: z.string(),
  modificationRecurringEvents: zu.stringToJSON().pipe(
    z.object({
      after: z.coerce.date(),
      events: z
        .object({
          title: z.string(),
          dayOfWeek: DayOfWeek,
          startTime: z.coerce.date(),
          endTime: z.coerce.date(),
        })
        .array(),
    }),
  ),
});
type ModifyRecurringEventRequest = z.infer<typeof ModifyRecurringEventRequest>;

const DeleteRecurringEventRequest = z.object({
  operationType: z.literal("deleteRecurringEvent"),
  userEmail: z.string(),
  deletionRecurringEvents: zu.stringToJSON().pipe(
    z.object({
      after: z.coerce.date(),
      dayOfWeeks: DayOfWeek.array(),
    }),
  ),
});
type DeleteRecurringEventRequest = z.infer<typeof DeleteRecurringEventRequest>;

type DeleteRecurringEventResponse = {
  responseCode: number;
  comment: string;
};

const ShiftChangeRequestSchema = z.union([
  RegisterEventRequest,
  ModifyAndDeleteEventRequest,
  ShowEventRequest,
  RegisterRecurringEventRequest,
  ModifyRecurringEventRequest,
  DeleteRecurringEventRequest,
]);

export const shiftChanger = (e: GoogleAppsScript.Events.DoPost) => {
  const parameter = ShiftChangeRequestSchema.parse(e.parameter);
  const operationType = parameter.operationType;
  const userEmail = parameter.userEmail;
  switch (operationType) {
    case "registerEvent": {
      registerEvents(userEmail, parameter.registrationEvents);
      break;
    }
    case "modifyAndDeleteEvent": {
      modifyEvents(parameter.modificationEvents, userEmail);
      deleteEvents(parameter.deletionEvents, userEmail);
      break;
    }
    case "showEvents": {
      const startDate = parameter.startDate;
      const eventEvent = showEvents(userEmail, startDate);

      return JSON.stringify(eventEvent);
    }
    case "registerRecurringEvent": {
      registerRecurringEvents(parameter, userEmail);
      break;
    }
    case "deleteRecurringEvent": {
      return JSON.stringify(deleteRecurringEvents(parameter, userEmail));
    }
    case "modifyRecurringEvent": {
      return JSON.stringify(modifyRecurringEvents(parameter, userEmail));
    }
  }
  return;
};

const registerEvents = (userEmail: string, registerInfos: Event[]) => {
  registerInfos.forEach((registerInfo) => {
    registerEvent(registerInfo, userEmail);
  });
};

const registerEvent = (eventInfo: Event, userEmail: string) => {
  const calendar = getCalendar();
  const [startDate, endDate] = [eventInfo.startTime, eventInfo.endTime];
  calendar.createEvent(eventInfo.title, startDate, endDate, { guests: userEmail });
};

const modifyEvents = (
  modifyInfos: {
    previousEvent: Event;
    newEvent: Event;
  }[],
  userEmail: string,
) => {
  const calendar = getCalendar();
  modifyInfos.forEach((eventInfo) => modifyEvent(eventInfo, calendar, userEmail));
};

const modifyEvent = (
  eventInfo: {
    previousEvent: Event;
    newEvent: Event;
  },
  calendar: GoogleAppsScript.Calendar.Calendar,
  userEmail: string,
) => {
  const [startDate, endDate] = [eventInfo.previousEvent.startTime, eventInfo.previousEvent.endTime];
  const newTitle = eventInfo.newEvent.title;
  const [newStartDate, newEndDate] = [eventInfo.newEvent.startTime, eventInfo.newEvent.endTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.setTime(newStartDate, newEndDate);
  event.setTitle(newTitle);
};

const deleteEvents = (deleteInfos: Event[], userEmail: string) => {
  const calendar = getCalendar();
  deleteInfos.forEach((eventInfo) => deleteEvent(eventInfo, calendar, userEmail));
};

const deleteEvent = (eventInfo: Event, calendar: GoogleAppsScript.Calendar.Calendar, userEmail: string) => {
  const [startDate, endDate] = [eventInfo.startTime, eventInfo.endTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.deleteEvent();
};

const showEvents = (userEmail: string, startDate: Date): Event[] => {
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

const registerRecurringEvents = (
  { registrationRecurringEvents: { after, events } }: RegisterRecurringEventRequest,
  userEmail: string,
) => {
  const calendar = getCalendar();

  events.forEach(({ title, startTime, endTime, dayOfWeek }) => {
    const recurrenceStartDate = getRecurrenceStartDate(after, dayOfWeek);
    const eventStartTime = mergeTimeToDate(recurrenceStartDate, startTime);
    const eventEndTime = mergeTimeToDate(recurrenceStartDate, endTime);
    const englishDayOfWeek = convertDayOfWeekJapaneseToEnglish(dayOfWeek);

    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(englishDayOfWeek);
    calendar.createEventSeries(title, eventStartTime, eventEndTime, recurrence, {
      guests: userEmail,
    });
  });
};

const modifyRecurringEvents = (
  { modificationRecurringEvents: { after, events } }: ModifyRecurringEventRequest,
  userEmail: string,
) => {
  const calendar = getCalendar();
  const calendarId = getConfig().CALENDAR_ID;
  const advancedCalendar = getAdvancedCalendar();

  //NOTE: 繰り返し予定を消去する機能
  const dayOfWeeks = events.map(({ dayOfWeek }) => dayOfWeek);
  const eventItems = dayOfWeeks
    .map((dayOfWeek) => {
      //NOTE: 仕様的にstartTimeの日付に最初の予定が指定されるため、指定された日付の前で一番近い指定曜日の日付に変更する
      const recurrenceEndDate = getRecurrenceEndDate(after, dayOfWeek);
      const events =
        advancedCalendar.list(calendarId, {
          timeMin: startOfDay(recurrenceEndDate).toISOString(),
          timeMax: endOfDay(recurrenceEndDate).toISOString(),
          singleEvents: true,
          orderBy: "startTime",
          maxResults: 1,
          q: userEmail,
        }).items ?? [];
      const recurringEventId = events[0].recurringEventId;
      return recurringEventId ? { recurringEventId, recurrenceEndDate } : undefined;
    })
    .filter(isNotUndefined);
  if (eventItems.length === 0) return { responseCode: 400, comment: "消去するイベントの取得に失敗しました" };

  const detailedEventItems = eventItems.map(({ recurringEventId, recurrenceEndDate }) => {
    const eventDetail = advancedCalendar.get(calendarId, recurringEventId);
    return { eventDetail, recurrenceEndDate, recurringEventId };
  });

  detailedEventItems.forEach(({ eventDetail, recurrenceEndDate, recurringEventId }) => {
    if (!eventDetail.start?.dateTime || !eventDetail.end?.dateTime) return;

    const untilTimeUTC = getEndOfDayFormattedAsUTCISO(recurrenceEndDate);
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
      recurrence: ["RRULE:FREQ=WEEKLY;UNTIL=" + untilTimeUTC],
    };
    advancedCalendar.update(data, calendarId, recurringEventId);
  });

  //NOTE: 繰り返し予定を登録する機能
  events.forEach(({ title, startTime, endTime, dayOfWeek }) => {
    const recurrenceStartDate = getRecurrenceStartDate(after, dayOfWeek);
    const eventStartTime = mergeTimeToDate(recurrenceStartDate, startTime);
    const eventEndTime = mergeTimeToDate(recurrenceStartDate, endTime);
    const englishDayOfWeek = convertDayOfWeekJapaneseToEnglish(dayOfWeek);

    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(englishDayOfWeek);
    calendar.createEventSeries(title, eventStartTime, eventEndTime, recurrence, {
      guests: userEmail,
    });
  });
  return { responseCode: 200, comment: "イベントの変更が成功しました" };
};

const deleteRecurringEvents = (
  { deletionRecurringEvents: { after, dayOfWeeks } }: DeleteRecurringEventRequest,
  userEmail: string,
): DeleteRecurringEventResponse => {
  const calendarId = getConfig().CALENDAR_ID;
  const advancedCalendar = getAdvancedCalendar();

  const eventItems = dayOfWeeks
    .map((dayOfWeek) => {
      //NOTE: 仕様的にstartTimeの日付に最初の予定が指定されるため、指定された日付の後で一番近い指定曜日の日付に変更する
      const recurrenceEndDate = getRecurrenceEndDate(after, dayOfWeek);
      const events =
        advancedCalendar.list(calendarId, {
          timeMin: startOfDay(recurrenceEndDate).toISOString(),
          timeMax: endOfDay(recurrenceEndDate).toISOString(),
          singleEvents: true,
          orderBy: "startTime",
          maxResults: 1,
          q: userEmail,
        }).items ?? [];
      const recurringEventId = events[0]?.recurringEventId;
      return recurringEventId ? { recurringEventId, recurrenceEndDate } : undefined;
    })
    .filter(isNotUndefined);
  if (eventItems.length === 0) return { responseCode: 400, comment: "消去するイベントの取得に失敗しました" };

  const detailedEventItems = eventItems.map(({ recurringEventId, recurrenceEndDate }) => {
    const eventDetail = advancedCalendar.get(calendarId, recurringEventId);
    return { eventDetail, recurrenceEndDate, recurringEventId };
  });

  detailedEventItems.forEach(({ eventDetail, recurrenceEndDate, recurringEventId }) => {
    if (!eventDetail.start?.dateTime || !eventDetail.end?.dateTime) return;

    const untilTimeUTC = getEndOfDayFormattedAsUTCISO(recurrenceEndDate);
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
      recurrence: ["RRULE:FREQ=WEEKLY;UNTIL=" + untilTimeUTC],
    };
    advancedCalendar.update(data, calendarId, recurringEventId);
  });

  return { responseCode: 200, comment: "イベントの消去が成功しました" };
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

const convertDayOfWeekJapaneseToEnglish = (dayOfWeek: DayOfWeek) => {
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

const convertDayOfWeekJapaneseToNumber = (dayOfWeek: DayOfWeek) => {
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

const getRecurrenceStartDate = (after: Date, dayOfWeek: DayOfWeek): Date => {
  const targetDayOfWeek = convertDayOfWeekJapaneseToNumber(dayOfWeek);
  if (after.getDay() === targetDayOfWeek) return after;
  const nextDate = nextDay(after, targetDayOfWeek);

  return nextDate;
};

const getRecurrenceEndDate = (after: Date, dayOfWeek: DayOfWeek): Date => {
  const targetDayOfWeek = convertDayOfWeekJapaneseToNumber(dayOfWeek);
  const previousDate = previousDay(after, targetDayOfWeek);

  return previousDate;
};

const isNotUndefined = <T>(value: T | undefined): value is T => {
  return value !== undefined;
};

const mergeTimeToDate = (date: Date, time: Date): Date => {
  return set(date, { hours: time.getHours(), minutes: time.getMinutes() });
};

const getEndOfDayFormattedAsUTCISO = (date: Date): string => {
  const endTime = endOfDay(date);
  const UTCTime = subHours(endTime, 9);
  return format(UTCTime, "yyyyMMdd'T'HHmmss'Z'");
};

const getAdvancedCalendar = () => {
  const advancedCalendar = Calendar.Events;
  if (advancedCalendar === undefined) throw new Error("カレンダーの取得に失敗しました");
  return advancedCalendar;
};
