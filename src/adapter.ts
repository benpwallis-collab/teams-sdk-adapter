import { BotFrameworkAdapter, TurnContext } from "botbuilder";

/**
 * Bot Framework ONLY reads these env vars for auth:
 *
 * MicrosoftAppId           (required)
 * MicrosoftAppPassword     (required)
 * MicrosoftAppType         (must be "SingleTenant")
 * MicrosoftAppTenantId     (required for SingleTenant)
 */

const {
  MicrosoftAppId,
  MicrosoftAppPassword,
  MicrosoftAppType,
  MicrosoftAppTenantId,
} = process.env;

if (!MicrosoftAppId || !MicrosoftAppPassword) {
  throw new Error("Missing MicrosoftAppId or MicrosoftAppPassword");
}

if (MicrosoftAppType !== "SingleTenant") {
  throw new Error(
    `MicrosoftAppType must be "SingleTenant", got "${MicrosoftAppType}"`
  );
}

if (!MicrosoftAppTenantId) {
  throw new Error("Missing MicrosoftAppTenantId for SingleTenant bot");
}

export const adapter = new BotFrameworkAdapter({
  appId: MicrosoftAppId,
  appPassword: MicrosoftAppPassword,
  // These two are REQUIRED for stable cross-tenant replies
  appType: "SingleTenant",
  tenantId: MicrosoftAppTenantId,
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

  // This may still 401 if auth is broken, so swallow failures
  try {
    await context.sendActivity("Something went wrong.");
  } catch {}
};
