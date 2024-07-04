import { GasWebClient as SlackClient } from "@hi-se/web-api";
import { set } from "date-fns";

import { getConfig } from "./config";
import { isBankHoliday } from "./date-utils";

const ANNOUNCE_HOUR = 9;

export function initNotifyDailyShift() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === notifyDailyShift.name)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger(notifyDailyShift.name)
    .timeBased()
    .atHour(ANNOUNCE_HOUR - 1)
    .nearMinute(30)
    .everyDays(1)
    .create();
}

export function notifyDailyShift() {
  const { CALENDAR_ID, SLACK_CHANNEL_TO_POST } = getConfig();
  const client = getSlackClient();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  const now = new Date();
  if (isBankHoliday(now)) return;
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
  const { SLACK_ACCESS_TOKEN } = getConfig();
  return new SlackClient(SLACK_ACCESS_TOKEN);
}
