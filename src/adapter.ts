import { BotFrameworkAdapter, TurnContext } from "botbuilder";

/**
 * Bot Framework authentication is driven ONLY by these env vars:
 *
 * MicrosoftAppId
 * MicrosoftAppPassword
 * MicrosoftAppTenantId
 * MicrosoftAppType=SingleTenant
 */

// --------------------------------------------------
// Startup diagnostics (DO NOT REMOVE)
// --------------------------------------------------
console.log("🔎 Adapter startup env check", {
  MicrosoftAppId: process.env.MicrosoftAppId ? "SET" : "MISSING",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword ? "SET" : "MISSING",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId ?? "(empty)",
  MicrosoftAppType: process.env.MicrosoftAppType ?? "(unset)",
});

// --------------------------------------------------
// Read env vars
// --------------------------------------------------
const appId = process.env.MicrosoftAppId;
const appPassword = process.env.MicrosoftAppPassword;

// --------------------------------------------------
// Hard fail with LOG (not throw)
// --------------------------------------------------
if (!appId || !appPassword) {
  console.error("❌ FATAL: Missing MicrosoftAppId or MicrosoftAppPassword");
  process.exit(1);
}

// --------------------------------------------------
// Adapter
// --------------------------------------------------
export const adapter = new BotFrameworkAdapter({
  appId,
  appPassword,
});

// --------------------------------------------------
// 🔍 PER-TURN DIAGNOSTICS (LOGGING ONLY)
// --------------------------------------------------
adapter.use(async (context, next) => {
  const activity = context.activity;

  console.log("📨 Incoming activity", {
    type: activity.type,
    channelId: activity.channelId,
    conversationType: activity.conversation?.conversationType,
    serviceUrl: activity.serviceUrl,
    tenantId:
      activity.channelData?.tenant?.id ??
      activity.conversation?.tenantId ??
      "(unknown)",
    fromId: activity.from?.id,
  });

  await next();
});

// --------------------------------------------------
// Global error handler
// --------------------------------------------------
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
          authorizationHeaderPresent: Boolean(
            err.request.headers?.authorization
          ),
        }
      : undefined,
    serviceUrlUsed: context.activity?.serviceUrl,
  });

  // Attempt to notify user (expected to fail on auth issues)
  try {
    await context.sendActivity("Something went wrong.");
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", {
      message: (sendErr as any)?.message,
      statusCode: (sendErr as any)?.statusCode,
    });
  }
};
