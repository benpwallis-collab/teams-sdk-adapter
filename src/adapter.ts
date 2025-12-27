import { BotFrameworkAdapter } from "botbuilder";

const appId = process.env.TEAMS_BOT_APP_ID;
const appPassword = process.env.TEAMS_BOT_APP_PASSWORD;

if (!appId || !appPassword) {
  throw new Error("Missing TEAMS_BOT_APP_ID or TEAMS_BOT_APP_PASSWORD");
}

/**
 * IMPORTANT:
 * Multi-tenant configuration is controlled via ENV VARS:
 *
 * MicrosoftAppType=MultiTenant
 * MicrosoftAppTenantId=common
 *
 * NOT via constructor options in botbuilder 4.23.x
 */
export const adapter = new BotFrameworkAdapter({
  appId,
  appPassword,
});

adapter.onTurnError = async (context, error) => {
  console.error("❌ onTurnError diagnostics:", {
    message: error.message,
    name: error.name,
    stack: error.stack,
    statusCode: (error as any)?.statusCode,
    details: (error as any)?.details,
    request: (error as any)?.request
      ? {
          method: (error as any).request.method,
          url: (error as any).request.url,
        }
      : undefined,
  });

  try {
    await context.sendActivity("Something went wrong.");
  } catch (err) {
    console.error("❌ Failed to send fallback message", err);
  }
};
