import { z } from "zod";

export const DayOfWeek = z
  .literal("月曜日")
  .or(z.literal("火曜日"))
  .or(z.literal("水曜日"))
  .or(z.literal("木曜日"))
  .or(z.literal("金曜日"));
export type DayOfWeek = z.infer<typeof DayOfWeek>;

// NOTE: z.object内でz.literal("").or(z.date())を使うと型推論がおかしくなるので、preprocessを使っている
export const DateOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), z.date().optional());

export const DayOfWeekOrEmptyString = z.preprocess((val) => (val === "" ? undefined : val), DayOfWeek.optional());

export const DateAfterNow = z.date().min(new Date(), { message: "過去の時間にシフト変更はできません" });

export const WorkingStyleOrEmptyString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("リモート").or(z.literal("出勤")).optional(),
);

export const OperationString = z.preprocess(
  (val) => (val === "" ? undefined : val),
  z.literal("時間変更").or(z.literal("消去")).or(z.literal("追加")).optional(),
);
