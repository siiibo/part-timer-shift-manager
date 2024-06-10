import { addWeeks, endOfDay, format, nextDay, previousDay, set, startOfDay, subHours } from "date-fns";
import { z } from "zod";

import { DayOfWeek } from "./common.schema";
import { getConfig } from "./config";

export const Event = z.object({
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type Event = z.infer<typeof Event>;

const RegisterEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("registerEvent"),
  userEmail: z.string(),
  events: Event.array(),
});
type RegisterEventRequest = z.infer<typeof RegisterEventRequest>;

const ModifyEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("modifyEvent"),
  userEmail: z.string(),
  events: z
    .object({
      previousEvent: Event,
      newEvent: Event,
    })
    .array(),
});
type ModifyEventRequest = z.infer<typeof ModifyEventRequest>;

const DeleteEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("deleteEvent"),
  userEmail: z.string(),
  events: Event.array(),
});
type DeleteEventRequest = z.infer<typeof DeleteEventRequest>;

const ShowEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("showEvents"),
  userEmail: z.string(),
  startDate: z.coerce.date(),
});
type ShowEventRequest = z.infer<typeof ShowEventRequest>;

const RegisterRecurringEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("registerRecurringEvent"),
  userEmail: z.string(),
  recurringInfo: z.object({
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
});
type RegisterRecurringEventRequest = z.infer<typeof RegisterRecurringEventRequest>;

const ModifyRecurringEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("modifyRecurringEvent"),
  userEmail: z.string(),
  recurringInfo: z.object({
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
});
type ModifyRecurringEventRequest = z.infer<typeof ModifyRecurringEventRequest>;

const DeleteRecurringEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("deleteRecurringEvent"),
  userEmail: z.string(),
  recurringInfo: z.object({
    after: z.coerce.date(),
    dayOfWeeks: DayOfWeek.array(),
  }),
});
type DeleteRecurringEventRequest = z.infer<typeof DeleteRecurringEventRequest>;

//NOTE: GASの仕様でレスポンスコードを返すことができないため、エラーメッセージを返す
export const RecurringEventResponse = z.object({
  error: z.string().optional(),
});
export type RecurringEventResponse = z.infer<typeof RecurringEventResponse>;

const ShiftChangeRequestSchema = z.union([
  RegisterEventRequest,
  ModifyEventRequest,
  DeleteEventRequest,
  ShowEventRequest,
  RegisterRecurringEventRequest,
  ModifyRecurringEventRequest,
  DeleteRecurringEventRequest,
]);
type ShiftChangeRequestSchema = z.infer<typeof ShiftChangeRequestSchema>;

export const doGet = () => {
  return ContentService.createTextOutput("ok");
};
export const doPost = (e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput => {
  const parameter = ShiftChangeRequestSchema.parse(JSON.parse(e.postData.contents));
  const response = shiftChanger(parameter) ?? "";
  return ContentService.createTextOutput(response).setMimeType(ContentService.MimeType.JSON);
};

export const shiftChanger = (parameter: ShiftChangeRequestSchema) => {
  const operationType = parameter.operationType;
  const userEmail = parameter.userEmail;
  switch (operationType) {
    case "registerEvent": {
      registerEvents(parameter.events, userEmail);
      break;
    }
    case "modifyEvent": {
      modifyEvents(parameter.events, userEmail);
      break;
    }
    case "deleteEvent": {
      deleteEvents(parameter.events, userEmail);
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
    case "modifyRecurringEvent": {
      return JSON.stringify(modifyRecurringEvents(parameter, userEmail));
    }
    case "deleteRecurringEvent": {
      return JSON.stringify(deleteRecurringEvents(parameter, userEmail));
    }
  }
  return;
};

const registerEvents = (registerInfos: Event[], userEmail: string) => {
  registerInfos.forEach((registerInfo) => {
    registerEvent(registerInfo, userEmail);
  });
};

const registerEvent = (eventInfo: Event, userEmail: string) => {
  const calendar = getCalendar(); //TODO: registerEventの引数にcalendarを追加する
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
    const startTime = new Date(event.getStartTime().getTime());
    const endTime = new Date(event.getEndTime().getTime());

    return { title, startTime, endTime };
  });
  return eventInfos;
};

const registerRecurringEvents = (
  { recurringInfo: { after, events } }: RegisterRecurringEventRequest,
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
  { recurringInfo: { after, events } }: ModifyRecurringEventRequest,
  userEmail: string,
): RecurringEventResponse => {
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
  if (eventItems.length === 0) return { error: "消去するイベントの取得に失敗しました" };

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

    const calendar = getCalendar();
    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(englishDayOfWeek);
    calendar.createEventSeries(title, eventStartTime, eventEndTime, recurrence, {
      guests: userEmail,
    });
  });
  return {}; //TODO: Result型導入時に削除
};

const deleteRecurringEvents = (
  { recurringInfo: { after, dayOfWeeks } }: DeleteRecurringEventRequest,
  userEmail: string,
): RecurringEventResponse => {
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
  if (eventItems.length === 0) return { error: "消去するイベントの取得に失敗しました" };

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
  return {}; //TODO: Result型導入時に削除
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
