import { BotFrameworkAdapter, TurnContext } from "botbuilder";

/**
 * Bot Framework auth is driven by ENV VARS ONLY.
 * We log them explicitly to prove what the SDK is using.
 */
const appId = process.env.MicrosoftAppId;
const appPassword = process.env.MicrosoftAppPassword;
const appTenantId = process.env.MicrosoftAppTenantId;
const appType = process.env.MicrosoftAppType;

console.log("🔐 Bot auth configuration at startup:", {
  MicrosoftAppId: appId ?? "MISSING",
  MicrosoftAppPassword: appPassword ? "SET" : "MISSING",
  MicrosoftAppTenantId: appTenantId ?? "(empty)",
  MicrosoftAppType: appType ?? "(default)",
});

if (!appId || !appPassword) {
  throw new Error("Missing MicrosoftAppId or MicrosoftAppPassword");
}

/**
 * IMPORTANT:
 * Do NOT pass appType / tenantId here.
 * BotFrameworkAdapter v4 reads them ONLY from env vars.
 */
export const adapter = new BotFrameworkAdapter({
  appId,
  appPassword,
});

/**
 * Global turn error handler with deep diagnostics.
 * This is where we see EXACTLY what fails on sendActivity.
 */
adapter.onTurnError = async (
  context: TurnContext,
  error: unknown
) => {
  const err = error as any;

  console.error("❌ onTurnError diagnostics:", {
    message: err?.message,
    name: err?.name,
    errorCode: err?.errorCode,
    statusCode: err?.statusCode,
    subError: err?.subError,
    correlationId: err?.correlationId,
    details: err?.details,
    request: err?.request
      ? {
          method: err.request.method,
          url: err.request.url,
          headers: err.request.headers
            ? {
                // redact auth, keep signal
                authorization: err.request.headers.authorization
                  ? "REDACTED"
                  : undefined,
                "x-ms-client-request-id":
                  err.request.headers["x-ms-client-request-id"],
              }
            : undefined,
        }
      : undefined,
  });

  /**
   * Attempt to notify user, but this will ALSO fail if auth is broken.
   * We swallow errors so we don't mask the real issue.
   */
  try {
    await context.sendActivity("Something went wrong.");
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", sendErr);
  }
};
