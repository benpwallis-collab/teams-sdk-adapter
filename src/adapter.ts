import { BotFrameworkAdapter } from "botbuilder";
import type { TurnContext } from "botbuilder-core";

/**
 * Bot Framework authentication is driven by ENV VARS ONLY:
 *
 * MicrosoftAppId
 * MicrosoftAppPassword
 * MicrosoftAppTenantId
 * MicrosoftAppType=SingleTenant
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

// --------------------------------------------------
// Error handling
// --------------------------------------------------
adapter.onTurnError = async (
  context: TurnContext,
  error: unknown
): Promise<void> => {
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
