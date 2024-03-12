import { addWeeks } from "date-fns";
import { z } from "zod";

import { getConfig } from "./config";

export const RegistrationInfo = z.object({
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type RegistrationInfo = z.infer<typeof RegistrationInfo>;

export const ModificationInfo = z.object({
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
  newTitle: z.string(),
  newStartTime: z.coerce.date(),
  newEndTime: z.coerce.date(),
});
export type ModificationInfo = z.infer<typeof ModificationInfo>;

export const DeletionInfo = z.object({
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type DeletionInfo = z.infer<typeof DeletionInfo>;

export const ShowEventInfo = z.object({
  title: z.string(),
  startTime: z.coerce.date(),
  endTime: z.coerce.date(),
});
export type ShowEventInfo = z.infer<typeof ShowEventInfo>;

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
      const registrationInfos = RegistrationInfo.array().parse(JSON.parse(e.parameter.registrationInfos));
      registration(userEmail, registrationInfos);
      break;
    }
    case "modificationAndDeletion": {
      const modificationInfos = ModificationInfo.array().parse(JSON.parse(e.parameter.modificationInfos));
      const deletionInfos = DeletionInfo.array().parse(JSON.parse(e.parameter.deletionInfos));
      modification(modificationInfos, userEmail);
      deletion(deletionInfos, userEmail);
      break;
    }
    case "showEvents": {
      const startDate = new Date(e.parameter.startDate);
      const showEventInfo = showEvents(userEmail, startDate);
      return JSON.stringify(showEventInfo);
    }
  }
  return;
};

const registration = (userEmail: string, registrationInfos: RegistrationInfo[]) => {
  registrationInfos.forEach((registrationInfo) => {
    registerEvent(registrationInfo, userEmail);
  });
};

const registerEvent = (registrationInfo: RegistrationInfo, userEmail: string) => {
  const calendar = getCalendar();
  const [startDate, endDate] = [registrationInfo.startTime, registrationInfo.endTime];
  calendar.createEvent(registrationInfo.title, startDate, endDate, { guests: userEmail });
};

const showEvents = (userEmail: string, startDate: Date): ShowEventInfo[] => {
  const endDate = addWeeks(startDate, 4);
  const calendar = getCalendar();
  const events = calendar.getEvents(startDate, endDate).filter((event) => isEventGuest(event, userEmail));
  const showEventInfos = events.map((event) => {
    const title = event.getTitle();
    const startTime = new Date(event.getStartTime().getTime());
    const endTime = new Date(event.getEndTime().getTime());

    return { title, startTime, endTime };
  });
  return showEventInfos;
};

const modification = (modificationInfos: ModificationInfo[], userEmail: string) => {
  const calendar = getCalendar();
  modificationInfos.forEach((modificationInfo) => modifyEvent(modificationInfo, calendar, userEmail));
};

const modifyEvent = (
  modificationInfo: ModificationInfo,
  calendar: GoogleAppsScript.Calendar.Calendar,
  userEmail: string,
) => {
  const [startDate, endDate] = [modificationInfo.startTime, modificationInfo.endTime];
  const newTitle = modificationInfo.newTitle;
  const [newStartDate, newEndDate] = [modificationInfo.newStartTime, modificationInfo.newEndTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.setTime(newStartDate, newEndDate);
  event.setTitle(newTitle);
};

const deletion = (deletionInfos: DeletionInfo[], userEmail: string) => {
  const calendar = getCalendar();
  deletionInfos.forEach((deletionInfo) => deleteEvent(deletionInfo, calendar, userEmail));
};

const deleteEvent = (deletionInfo: DeletionInfo, calendar: GoogleAppsScript.Calendar.Calendar, userEmail: string) => {
  const [startDate, endDate] = [deletionInfo.startTime, deletionInfo.endTime];
  const event = calendar.getEvents(startDate, endDate).find((event) => isEventGuest(event, userEmail));
  if (!event) return;
  event.deleteEvent();
};
