import { addDays, addWeeks } from "date-fns";
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

const RegistrationRecurringEvent = z.object({
  dayOfWeek: dayOfWeek,
  startOrEndDate: z.coerce.date(),
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
type RegistrationRecurringEvent = z.infer<typeof RegistrationRecurringEvent>;

const DeletionRecurringEvent = z.object({
  endDate: z.coerce.date(),
  dayOfWeek: dayOfWeek,
});
type DeletionRecurringEvent = z.infer<typeof DeletionRecurringEvent>;

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
      const registrationRecurringEvents = RegistrationRecurringEvent.array().parse(
        JSON.parse(e.parameter.recurringEventModification),
      );

      registerRecurringEvent(registrationRecurringEvents, userEmail);
      break;
    }
    case "deleteRecurringEvent": {
      const deletionRecurringEvents = DeletionRecurringEvent.array().parse(
        JSON.parse(e.parameter.recurringEventDeletion),
      );
      deleteRecurringEvent(deletionRecurringEvents);
      break;
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

const registerRecurringEvent = (registrationRecurringEvents: RegistrationRecurringEvent[], userEmail: string) => {
  const calendar = getCalendar();
  registrationRecurringEvents.forEach((event) => {
    const dayOfWeek = convertJapaneseToEnglishDayOfWeek(event.dayOfWeek);
    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(dayOfWeek);
    calendar.createEventSeries(event.title, event.startTime, event.endTime, recurrence, {
      guests: userEmail,
    });
  });
};
const deleteRecurringEvent = (deletionRecurringEvents: DeletionRecurringEvent[]) => {
  const calendar = getCalendar();
  deletionRecurringEvents.forEach((event) => {
    const eventId = calendar.getEvents(event.endDate, addDays(event.endDate, 1))[0].getId();
    console.log(eventId);
    const url =
      "https://www.googleapis.com/calendar/v3/calendars/" + encodeURIComponent(calendar.getId()) + "/events/" + eventId;

    const endDate = event.endDate;
    const data = {
      end: {
        date: endDate.toISOString().substring(0, 10), // YYYY-MM-DD 形式
      },
      recurrence: [
        "RRULE:FREQ=DAILY;UNTIL=" + endDate.toISOString().replace(/-|:|\.\d+Z$/g, "") + "Z", // UNTIL=YYYYMMDDT000000Z 形式
      ],
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "patch",
      payload: JSON.stringify(data),
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken(),
      },
    };
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(response.getContentText());
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
