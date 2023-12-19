import { format } from "date-fns";
import { addWeeks } from "date-fns";

import { getConfig } from "./config";

export type EventInfo = { title: string; date: Date; startTime: Date; endTime: Date };

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
      const registrationInfos = JSON.parse(e.parameter.registrationInfos);
      registration(userEmail, registrationInfos);
      break;
    }
    case "modificationAndDeletion": {
      const modificationInfos = JSON.parse(e.parameter.modificationInfos);
      const deletionInfos = JSON.parse(e.parameter.deletionInfos);
      modification(modificationInfos, userEmail);
      deletion(deletionInfos, userEmail);
      break;
    }
    case "showEvents": {
      const startDate = new Date(e.parameter.startDate);
      const eventInfos = showEvents(userEmail, startDate);
      const formattedEvents = eventInfos.map((event) => {
        const title = event.title;
        const date = Utilities.formatDate(event.date, "JST", "MM/dd");
        const startTime = Utilities.formatDate(event.startTime, "JST", "HH:mm");
        const endTime = Utilities.formatDate(event.endTime, "JST", "HH:mm");
        return { title, date, startTime, endTime };
      });
      return JSON.stringify(formattedEvents);
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
  const [startDate, endDate] = getStartEndDate(eventInfo);
  calendar.createEvent(eventInfo.title, startDate, endDate, { guests: userEmail });
};

const showEvents = (userEmail: string, startDate: Date): EventInfo[] => {
  const endDate = addWeeks(startDate, 4);
  const calendar = getCalendar();
  const events = calendar.getEvents(startDate, endDate).filter((event) => isEventGuest(event, userEmail));
  const eventInfos = events.map((event) => {
    const title = event.getTitle();
    const date =event.getStartTime() as Date;
    const startTime = event.getStartTime() as Date;
    const endTime = event.getEndTime() as Date;

    return { title, date, startTime, endTime };
  });
  return eventInfos;
};

const modification = (
  modificationInfos: {
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  }[],
  userEmail: string
) => {
  const calendar = getCalendar();
  modificationInfos.forEach((eventInfo) => modifyEvent(eventInfo, calendar, userEmail));
};

const modifyEvent = (
  eventInfo: {
    previousEventInfo: EventInfo;
    newEventInfo: EventInfo;
  },
  calendar: GoogleAppsScript.Calendar.Calendar,
  userEmail: string
) => {
  const [startDate, endDate] = getStartEndDate(eventInfo.previousEventInfo);

  const newTitle = eventInfo.newEventInfo.title;

  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.setTime(startDate, endDate);
  event.setTitle(newTitle);
};

const getStartEndDate = ({ date,startTime, endTime }: EventInfo): [Date, Date] => {
  startTime.setFullYear(date.getFullYear());
  startTime.setMonth(date.getMonth());
  startTime.setDate(date.getDate());

  endTime.setFullYear(date.getFullYear());
  endTime.setMonth(date.getMonth());
  endTime.setDate(date.getDate());
    return [startTime, endTime];
}

const deletion = (deletionInfos: EventInfo[], userEmail: string) => {
  const calendar = getCalendar();
  deletionInfos.forEach((eventInfo) => deleteEvent(eventInfo, calendar, userEmail));
};

const deleteEvent = (eventInfo: EventInfo, calendar: GoogleAppsScript.Calendar.Calendar, userEmail: string) => {
  const [startDate, endDate] = getStartEndDate(eventInfo);
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.deleteEvent();
};
