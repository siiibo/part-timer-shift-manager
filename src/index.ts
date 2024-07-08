import { deleteHolidayShift, initDeleteHolidayShift } from "./autoDeleteHolidayEvent";
import { insertModificationAndDeletionSheet } from "./ModificationAndDeletionSheet";
import { initNotifyDailyShift, notifyDailyShift } from "./notify-daily-shift";
import { insertRecurringEventSheet } from "./RecurringEventSheet";
import { insertRegistrationSheet } from "./RegistrationSheet";
import {
  callModificationAndDeletion,
  callRecurringEvent,
  callRegistration,
  callShowEvents,
  initShiftChanger,
  onOpen,
  onOpenForDev,
} from "./shift-changer";
import { doGet, doPost } from "./shift-changer-api";
/**
 * @file GASエディタから実行できる関数を定義する
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
declare const global: any;
global.initShiftChanger = initShiftChanger;
global.initNotifyDailyShift = initNotifyDailyShift;
global.notifyDailyShift = notifyDailyShift;
global.doPost = doPost;
global.onOpen = onOpen;
global.doGet = doGet;
global.onOpenForDev = onOpenForDev;
global.callRegistration = callRegistration;
global.callShowEvents = callShowEvents;
global.callModificationAndDeletion = callModificationAndDeletion;
global.insertRegistrationSheet = insertRegistrationSheet;
global.insertModificationAndDeletionSheet = insertModificationAndDeletionSheet;
global.insertRecurringEventSheet = insertRecurringEventSheet;
global.callRecurringEvent = callRecurringEvent;
global.initDeleteHolidayShift = initDeleteHolidayShift;
global.deleteHolidayShift = deleteHolidayShift;
