import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { isWeekend, set } from "date-fns";

import { getConfig } from "./config";

const ANNOUNCE_HOUR = 9;

export function initNotifyDailyShift() {
  const targetFunction = notifyDailyShift;

  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === targetFunction.name) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(targetFunction.name)
    .timeBased()
    .atHour(ANNOUNCE_HOUR - 1)
    .everyDays(1)
    .create();
}

export function notifyDailyShift() {
  const { CALENDAR_ID, SLACK_CHANNEL_TO_POST } = getConfig();
  const client = getSlackClient();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  const now = new Date();
  if (isWeekend(now) || isHoliday(now)) return;
  if (!checkTime(now)) throw new Error(`設定時刻に誤りがあります.\nANNOUNCE_HOUR: ${ANNOUNCE_HOUR}\nnow: ${now}`);

  const targetDate = new Date();
  const announceTime = set(targetDate, { hours: ANNOUNCE_HOUR, minutes: 0, seconds: 0, milliseconds: 0 });
  const dailyShifts = calendar.getEventsForDay(targetDate);
  const notificationString = getNotificationString(dailyShifts);

  client.chat.scheduleMessage({
    channel: SLACK_CHANNEL_TO_POST,
    post_at: getUnixTimeStampString(announceTime),
    text: notificationString,
  });
}

function getNotificationString(events: GoogleAppsScript.Calendar.CalendarEvent[]): string {
  return !events.length
    ? "今日の予定はありません"
    : events.map(getNotificationStringForEvent).join("\n") +
        "\n\n" +
        ":calendar: 勤務開始時に<https://calendar.google.com/calendar|カレンダー>に予定が入っていないか確認しましょう！";
}

function getNotificationStringForEvent(event: GoogleAppsScript.Calendar.CalendarEvent): string {
  const title = event.getTitle();
  const startTime = Utilities.formatDate(event.getStartTime(), "Asia/Tokyo", "HH:mm");
  const endTime = Utilities.formatDate(event.getEndTime(), "Asia/Tokyo", "HH:mm");
  return `${title}  ${startTime} 〜 ${endTime}`;
}

function checkTime(target: Date) {
  return target.getHours() === ANNOUNCE_HOUR - 1;
}

function getUnixTimeStampString(date: Date): string {
  return Math.floor(date.getTime() / 1000).toFixed();
}

function getSlackClient() {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");
  if (!token) throw new Error("SLACK_TOKEN is not set");
  return new SlackClient(token);
}

function isHoliday(day: Date): boolean {
  const calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
  const calendar = CalendarApp.getCalendarById(calendarId);
  const holidayEvents = calendar.getEventsForDay(day);
  return holidayEvents.length > 0;
}
