import { BotFrameworkAdapter, TurnContext } from "botbuilder";

/**
 * Bot Framework authentication is driven ONLY by these env vars:
 *
 * MicrosoftAppId
 * MicrosoftAppPassword
 * MicrosoftAppTenantId   (required for SingleTenant)
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
        }
      : undefined,
  });

  // Attempt to notify user (will fail on auth issues, that's fine)
  try {
    await context.sendActivity("Something went wrong.");
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", sendErr);
  }
};
