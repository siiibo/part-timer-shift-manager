import {
  callModificationAndDeletion,
  callRegistration,
  callShowEvents,
  doPost,
  insertModificationAndDeletionSheet,
  insertRegistrationSheet,
  onOpen,
} from "./shift-changer";

/**
 * @file GASエディタから実行できる関数を定義する
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
declare const global: any;
global.doPost = doPost;
global.onOpen = onOpen;
global.callRegistration = callRegistration;
global.callShowEvents = callShowEvents;
global.callModificationAndDeletion = callModificationAndDeletion;
global.insertRegistrationSheet = insertRegistrationSheet;
global.insertModificationAndDeletionSheet = insertModificationAndDeletionSheet;
