import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
} from "botbuilder";

/**
 * Startup diagnostics
 */
console.log("🔐 Bot auth configuration at startup:", {
  MicrosoftAppId: process.env.MicrosoftAppId ?? "MISSING",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword ? "SET" : "MISSING",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId ?? "(empty)",
  MicrosoftAppType: process.env.MicrosoftAppType ?? "(unset)",
});

// Hard fail early if auth is incomplete
if (
  !process.env.MicrosoftAppId ||
  !process.env.MicrosoftAppPassword ||
  !process.env.MicrosoftAppTenantId
) {
  throw new Error(
    "Missing MicrosoftAppId, MicrosoftAppPassword, or MicrosoftAppTenantId"
  );
}

/**
 * Bot Framework authentication
 * This is REQUIRED for SingleTenant bots (2025+)
 */
const botFrameworkAuthentication =
  new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: process.env.MicrosoftAppId,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword,
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId,
    MicrosoftAppType: "SingleTenant",
  });

/**
 * CloudAdapter
 */
export const adapter = new CloudAdapter(botFrameworkAuthentication);

/**
 * Global error handler
 */
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
    serviceUrl: context.activity?.serviceUrl,
  });

  try {
    await context.sendActivity("Something went wrong.");
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", {
      message: (sendErr as any)?.message,
      statusCode: (sendErr as any)?.statusCode,
    });
  }
};
