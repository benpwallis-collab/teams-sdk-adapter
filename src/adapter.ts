import { BotFrameworkAdapter } from "botbuilder";

const appId = process.env.TEAMS_BOT_APP_ID;
const appPassword = process.env.TEAMS_BOT_APP_PASSWORD;

if (!appId || !appPassword) {
  throw new Error("Missing TEAMS_BOT_APP_ID or TEAMS_BOT_APP_PASSWORD");
}

export const adapter = new BotFrameworkAdapter({
  appId,
  appPassword,
});

adapter.onTurnError = async (context, error) => {
  console.error("❌ onTurnError:", error);
  try {
    await context.sendActivity("Something went wrong.");
  } catch {}
};
