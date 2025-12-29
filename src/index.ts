import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  TurnContext,
} from "botbuilder";

/**
 * Bot Framework auth is driven by ENV VARS.
 * We log them explicitly to prove runtime configuration.
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

if (!appId || !appPassword || !appTenantId) {
  throw new Error(
    "Missing MicrosoftAppId, MicrosoftAppPassword, or MicrosoftAppTenantId"
  );
}

/**
 * Credential factory (REQUIRED for SingleTenant)
 */
const credentialsFactory =
  new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: appId,
    MicrosoftAppPassword: appPassword,
    MicrosoftAppTenantId: appTenantId,
    MicrosoftAppType: "SingleTenant",
  });

/**
 * CloudAdapter (required for 2025 auth model)
 */
export const adapter = new CloudAdapter(credentialsFactory);

/**
 * Global turn error handler with deep diagnostics.
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
