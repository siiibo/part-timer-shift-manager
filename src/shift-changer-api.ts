import { addWeeks, endOfDay, format, nextDay, startOfDay, subDays, subHours, subWeeks } from "date-fns";
import { type Result, err, ok } from "neverthrow";
import { z } from "zod";

import { DayOfWeek } from "./common.schema";
import { getConfig } from "./config";
import { mergeTimeToDate } from "./date-utils";

export const Event = z.object({
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type Event = z.infer<typeof Event>;

export const RegisterEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("registerEvent"),
  userEmail: z.string(),
  events: Event.array(),
});
export type RegisterEventRequest = z.infer<typeof RegisterEventRequest>;

export const ModifyEventRequest = z.object({
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
export type ModifyEventRequest = z.infer<typeof ModifyEventRequest>;

export const DeleteEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("deleteEvent"),
  userEmail: z.string(),
  events: Event.array(),
});
export type DeleteEventRequest = z.infer<typeof DeleteEventRequest>;

export const ShowEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("showEvents"),
  userEmail: z.string(),
  startDate: z.coerce.date(),
});
export type ShowEventRequest = z.infer<typeof ShowEventRequest>;

export const RegisterRecurringEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("registerRecurringEvent"),
  userEmail: z.string(),
  recurringInfo: z.object({
    newShiftStartDate: z.coerce.date(),
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
export type RegisterRecurringEventRequest = z.infer<typeof RegisterRecurringEventRequest>;

export const ModifyRecurringEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("modifyRecurringEvent"),
  userEmail: z.string(),
  recurringInfo: z.object({
    newShiftStartDate: z.coerce.date(),
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
export type ModifyRecurringEventRequest = z.infer<typeof ModifyRecurringEventRequest>;

export const DeleteRecurringEventRequest = z.object({
  apiId: z.literal("shift-changer"),
  operationType: z.literal("deleteRecurringEvent"),
  userEmail: z.string(),
  recurringInfo: z.object({
    newShiftStartDate: z.coerce.date(),
    dayOfWeeks: DayOfWeek.array(),
  }),
});
export type DeleteRecurringEventRequest = z.infer<typeof DeleteRecurringEventRequest>;

//NOTE: GASの仕様でレスポンスコードを返すことができないため、エラーメッセージを返す
export const APIResponse = z.union([
  z.object({
    error: z.string(),
  }),
  z.object({
    events: Event.array(),
  }),
]);
export type APIResponse = z.infer<typeof APIResponse>;

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
  const result = shiftChanger(parameter);
  return result
    .match(
      (maybeEvent) =>
        // NOTE: 2024/06時点ではshowEvents,modifyRecurringEvent,deleteRecurringEventがeventsを返す
        ContentService.createTextOutput(JSON.stringify({ events: maybeEvent === "成功" ? [] : maybeEvent })),
      (error) => ContentService.createTextOutput(JSON.stringify({ error })),
    )
    .setMimeType(ContentService.MimeType.JSON);
};

export const shiftChanger = (parameter: ShiftChangeRequestSchema): Result<Event[] | string | never, string> => {
  const operationType = parameter.operationType;
  const userEmail = parameter.userEmail;
  switch (operationType) {
    case "registerEvent": {
      return registerEvents(parameter.events, userEmail);
    }
    case "modifyEvent": {
      return modifyEvents(parameter.events, userEmail);
    }
    case "deleteEvent": {
      return deleteEvents(parameter.events, userEmail);
    }
    case "showEvents": {
      return showEvents(userEmail, parameter.startDate);
    }
    case "registerRecurringEvent": {
      return registerRecurringEvents(parameter, userEmail);
    }
    case "modifyRecurringEvent": {
      return modifyRecurringEvents(parameter, userEmail);
    }
    case "deleteRecurringEvent": {
      return deleteRecurringEvents(parameter, userEmail);
    }
  }
};

const registerEvents = (registerInfos: Event[], userEmail: string): Result<string, never> => {
  registerInfos.forEach((registerInfo) => {
    registerEvent(registerInfo, userEmail);
  });
  return ok("成功");
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
): Result<string, never> => {
  const calendar = getCalendar();
  modifyInfos.forEach((eventInfo) => modifyEvent(eventInfo, calendar, userEmail));
  return ok("成功");
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
  if (!event) {
    return;
  }
  event.setTime(newStartDate, newEndDate);
  event.setTitle(newTitle);
};

const deleteEvents = (deleteInfos: Event[], userEmail: string): Result<string, never> => {
  const calendar = getCalendar();
  deleteInfos.forEach((eventInfo) => deleteEvent(eventInfo, calendar, userEmail));
  return ok("成功");
};

const deleteEvent = (eventInfo: Event, calendar: GoogleAppsScript.Calendar.Calendar, userEmail: string) => {
  const [startDate, endDate] = [eventInfo.startTime, eventInfo.endTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) {
    return;
  }
  event.deleteEvent();
};

const showEvents = (userEmail: string, startDate: Date): Result<Event[], never> => {
  const endDate = addWeeks(startDate, 4);
  const calendar = getCalendar();
  const events = calendar.getEvents(startDate, endDate).filter((event) => isEventGuest(event, userEmail));
  const eventInfos = events.map((event) => {
    const title = event.getTitle();
    const startTime = new Date(event.getStartTime().getTime());
    const endTime = new Date(event.getEndTime().getTime());

    return { title, startTime, endTime };
  });
  return ok(eventInfos);
};

const registerRecurringEvents = (
  { recurringInfo: { newShiftStartDate, events } }: RegisterRecurringEventRequest,
  userEmail: string,
): Result<string, never> => {
  const calendar = getCalendar();

  events.forEach(({ title, startTime, endTime, dayOfWeek }) => {
    const recurrenceStartDate = getRecurrenceStartDate(newShiftStartDate, dayOfWeek);
    const eventStartTime = mergeTimeToDate(recurrenceStartDate, startTime);
    const eventEndTime = mergeTimeToDate(recurrenceStartDate, endTime);
    const englishDayOfWeek = convertDayOfWeekJapaneseToEnglish(dayOfWeek);

    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(englishDayOfWeek);
    calendar.createEventSeries(title, eventStartTime, eventEndTime, recurrence, {
      guests: userEmail,
    });
  });
  return ok("成功");
};

const modifyRecurringEvents = (
  { recurringInfo: { newShiftStartDate, events } }: ModifyRecurringEventRequest,
  userEmail: string,
): Result<Event[], string> => {
  const dayOfWeeks = events.map(({ dayOfWeek }) => dayOfWeek);
  return deleteRecurringEvents(
    {
      apiId: "shift-changer",
      operationType: "deleteRecurringEvent",
      userEmail,
      recurringInfo: { newShiftStartDate, dayOfWeeks },
    },
    userEmail,
  )
    .andThen((beforeEvents) => {
      registerRecurringEvents(
        {
          apiId: "shift-changer",
          operationType: "registerRecurringEvent",
          userEmail,
          recurringInfo: { newShiftStartDate, events },
        },
        userEmail,
      );
      return ok(beforeEvents);
    })
    .orElse((error) => {
      return err(error);
    });
};

const deleteRecurringEvents = (
  { recurringInfo: { newShiftStartDate, dayOfWeeks } }: DeleteRecurringEventRequest,
  userEmail: string,
): Result<Event[], string> => {
  const calendarId = getConfig().CALENDAR_ID;
  const advancedCalendar = getAdvancedCalendar();

  const events = getCandidateEventsOfOldShiftEnd(advancedCalendar, calendarId, newShiftStartDate, userEmail);

  const recurrenceEndEventIdsResult = getRecurrenceEndEventIds(events, dayOfWeeks);

  if (recurrenceEndEventIdsResult.isErr()) {
    return err(recurrenceEndEventIdsResult.error);
  }

  const detailedEvents = recurrenceEndEventIdsResult.value.map((recurringEventId) => {
    const eventDetail = advancedCalendar.get(calendarId, recurringEventId);
    //TODO: eventDetailの型をzodで定義してパースする
    return { eventDetail, recurringEventId };
  });

  const deleteEvents = detailedEvents
    .map(({ eventDetail, recurringEventId }) => {
      if (!(eventDetail.start?.dateTime && eventDetail.end?.dateTime && eventDetail.summary)) {
        //NOTE: この箇所でundefinedが発生する場合はないはず
        return;
      }
      const untilTimeUTC = getEndOfDayFormattedAsUTCISO(newShiftStartDate);
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
        recurrence: [`RRULE:FREQ=WEEKLY;UNTIL=${untilTimeUTC}`],
      };
      advancedCalendar.update(data, calendarId, recurringEventId);
      return {
        title: eventDetail.summary,
        startTime: new Date(eventDetail.start.dateTime),
        endTime: new Date(eventDetail.end.dateTime),
      };
    })
    .filter(isNotUndefined);
  return ok(deleteEvents);
};

function getCandidateEventsOfOldShiftEnd(
  advancedCalendar: GoogleAppsScript.Calendar.Collection.EventsCollection,
  calendarId: string,
  newShiftStartDate: Date,
  userEmail: string,
) {
  return (
    advancedCalendar.list(calendarId, {
      // NOTE: 1週間だと祝日等が被った際に予定が取得できない場合があるので、余裕を持って4週間分取得している
      timeMin: startOfDay(subWeeks(newShiftStartDate, 4)).toISOString(),
      // NOTE: 旧シフトの最終日候補を取得するため、timeMaxにはnewShiftStartDateを含めない
      timeMax: endOfDay(subDays(newShiftStartDate, 1)).toISOString(),
      singleEvents: true,
      orderBy: "startTime",
      q: userEmail,
    }).items ?? []
  );
}

const getRecurrenceEndEventIds = (
  events: GoogleAppsScript.Calendar.Schema.Event[],
  dayOfWeeks: DayOfWeek[],
): Result<string[], string> => {
  const recurrenceEndEventIds = dayOfWeeks.map((dayOfWeek) => {
    const targetDayOfWeek = convertDayOfWeekJapaneseToNumber(dayOfWeek);

    const sortedEventsByDateDescending = events.sort((a, b) => {
      const dateA = a.start?.dateTime ?? "";
      const dateB = b.start?.dateTime ?? "";
      return dateB.localeCompare(dateA);
    });

    // NOTE: 日付の降順でsortされたeventsから、指定した曜日の最初のrecurringEventIdを取得
    const event = sortedEventsByDateDescending.find((event) => {
      const eventDayOfWeek = event.start?.dateTime ? new Date(event.start.dateTime).getDay() : undefined;
      return eventDayOfWeek !== undefined && targetDayOfWeek === eventDayOfWeek && event.recurringEventId !== undefined;
    });
    return event?.recurringEventId;
  });

  //NOTE: 一つでもundefinedがある場合はerrを返す
  if (
    recurrenceEndEventIds.length === 0 ||
    recurrenceEndEventIds.some((recurrenceEndEventId) => recurrenceEndEventId === undefined)
  ) {
    return err("削除対象の予定が見つかりませんでした");
  }

  return ok(recurrenceEndEventIds.filter(isNotUndefined));
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

const getRecurrenceStartDate = (newShiftStartDate: Date, dayOfWeek: DayOfWeek): Date => {
  const targetDayOfWeek = convertDayOfWeekJapaneseToNumber(dayOfWeek);
  if (newShiftStartDate.getDay() === targetDayOfWeek) {
    return newShiftStartDate;
  }
  const nextDate = nextDay(newShiftStartDate, targetDayOfWeek);

  return nextDate;
};

const isNotUndefined = <T>(value: T | undefined): value is T => {
  return value !== undefined;
};

const getEndOfDayFormattedAsUTCISO = (date: Date): string => {
  const endTime = endOfDay(date);
  const UTCTime = subHours(endTime, 9);
  return format(UTCTime, "yyyyMMdd'T'HHmmss'Z'");
};

const getAdvancedCalendar = () => {
  const advancedCalendar = Calendar.Events;
  if (advancedCalendar === undefined) {
    throw new Error("カレンダーの取得に失敗しました");
  }
  return advancedCalendar;
};
