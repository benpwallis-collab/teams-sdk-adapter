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

adapter.onTurnError = async (context: any, error: any) => {
  console.error("❌ onTurnError diagnostics:", {
    message: error?.message,
    name: error?.name,
    code: error?.code,
    statusCode: error?.statusCode,
    details: error?.details,
    body: error?.response?.body,
    request: {
      method: error?.request?.method,
      url: error?.request?.url,
    },
  });

  try {
    await context.sendActivity(
      "Sorry, something went wrong while processing your message."
    );
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", sendErr);
  }
};
