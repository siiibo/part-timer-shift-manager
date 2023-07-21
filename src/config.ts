import { z } from "zod";

const ConfigSchema = z.object({
  API_URL: z.string(),
  CALENDAR_ID: z.string(),
  SLACK_ACCESS_TOKEN: z.string(),
  SLACK_CHANNEL_TO_POST: z.string(),
  JOB_SHEET_URL: z.string(),
  HR_MANAGER_SLACK_ID: z.string(),
});

export type Config = z.infer<typeof ConfigSchema>;

let config: Config;

export const getConfig = () => {
  if (!config) {
    const props = PropertiesService.getScriptProperties().getProperties();
    config = ConfigSchema.parse(props);
  }
  return config;
};
