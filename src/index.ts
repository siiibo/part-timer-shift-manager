import {
  callModificationAndDeletion,
  callShowEvents,
  insertModificationAndDeletionSheet,
} from "./ModificationAndDeletionSheet";
import { initNotifyDailyShift, notifyDailyShift } from "./notify-daily-shift";
import { callRegistration, insertRegistrationSheet } from "./RegistrationSheet";
import { doPost, initShiftChanger, onOpen, onOpenForDev } from "./shift-changer";
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
global.onOpenForDev = onOpenForDev;
global.callRegistration = callRegistration;
global.callShowEvents = callShowEvents;
global.callModificationAndDeletion = callModificationAndDeletion;
global.insertRegistrationSheet = insertRegistrationSheet;
global.insertModificationAndDeletionSheet = insertModificationAndDeletionSheet;
