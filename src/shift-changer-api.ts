import { addWeeks } from "date-fns";

import { getConfig } from "./config";

export type EventInfo = { title: string; date: string; startTime: string; endTime: string };

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
      return JSON.stringify(eventInfos);
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
    const date = Utilities.formatDate(event.getStartTime(), "JST", "MM/dd");
    const startTime = Utilities.formatDate(event.getStartTime(), "JST", "HH:mm");
    const endTime = Utilities.formatDate(event.getEndTime(), "JST", "HH:mm");

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

  const [newStartDate, newEndDate] = getStartEndDate(eventInfo.newEventInfo);

  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.setTime(newStartDate, newEndDate);
  event.setTitle(newTitle);
};

const getStartEndDate = ({ date, startTime, endTime }: EventInfo): [Date, Date] => {
  const startDate = new Date(`${date} ${startTime}`);
  const endDate = new Date(`${date} ${endTime}`);
  return [startDate, endDate];
};
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
