import { initNotifyDailyShift } from "./notify-daily-shift";
import {
  callModificationAndDeletion,
  callRegistration,
  callShowEvents,
  doPost,
  initShiftChanger,
  insertModificationAndDeletionSheet,
  insertRegistrationSheet,
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
global.doPost = doPost;
global.onOpen = onOpen;
global.onOpenForDev = onOpenForDev;
global.callRegistration = callRegistration;
global.callShowEvents = callShowEvents;
global.callModificationAndDeletion = callModificationAndDeletion;
global.insertRegistrationSheet = insertRegistrationSheet;
global.insertModificationAndDeletionSheet = insertModificationAndDeletionSheet;
