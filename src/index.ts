import { insertModificationAndDeletionSheet } from "./ModificationAndDeletionSheet";
import { initNotifyDailyShift, notifyDailyShift } from "./notify-daily-shift";
import { insertRecurringEventSheet } from "./RecurringShiftSheet";
import { insertRegistrationSheet } from "./RegistrationSheet";
import {
  callModificationAndDeletion,
  callRecurringEvent,
  callRegistration,
  callShowEvents,
  doGet,
  doPost,
  initShiftChanger,
  onOpen,
  onOpenForDev,
} from "./shift-changer";
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
