import { z } from "zod";

export const DayOfWeek = z
  .literal("月曜日")
  .or(z.literal("火曜日"))
  .or(z.literal("水曜日"))
  .or(z.literal("木曜日"))
  .or(z.literal("金曜日"));
export type DayOfWeek = z.infer<typeof DayOfWeek>;
