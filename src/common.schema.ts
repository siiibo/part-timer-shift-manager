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
  z.literal("出社").or(z.literal("リモート")).optional(),
);

//NOTE: min(1)で空文字を許容しないようにしている
export const CommentStringOrError = z.string().min(1, { message: "コメントを記入してください" });
