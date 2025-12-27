import { BotFrameworkAdapter, TurnContext } from "botbuilder";

/**
 * IMPORTANT:
 * Bot Framework ONLY reads these env vars for auth:
 *
 * MicrosoftAppId
 * MicrosoftAppPassword
 * MicrosoftAppType
 * MicrosoftAppTenantId
 */
const appId = process.env.MicrosoftAppId;
const appPassword = process.env.MicrosoftAppPassword;

if (!appId || !appPassword) {
  throw new Error("Missing MicrosoftAppId or MicrosoftAppPassword");
}

export const adapter = new BotFrameworkAdapter({
  appId,
  appPassword,
});

adapter.onTurnError = async (
  context: TurnContext,
  error: unknown
) => {
  const err = error as any;

  console.error("❌ onTurnError diagnostics:", {
    message: err?.message,
    name: err?.name,
    statusCode: err?.statusCode,
    details: err?.details,
    request: err?.request
      ? {
          method: err.request.method,
          url: err.request.url,
        }
      : undefined,
  });

  try {
    await context.sendActivity("Something went wrong.");
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", sendErr);
  }
};
