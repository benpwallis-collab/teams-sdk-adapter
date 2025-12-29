import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  TurnContext,
} from "botbuilder";

// --------------------------------------------------
// Credentials (SingleTenant compliant)
// --------------------------------------------------
const credentialsFactory =
  new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: process.env.MicrosoftAppId!,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword!,
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId!,
    MicrosoftAppType: "SingleTenant",
  });

// --------------------------------------------------
// Adapter
// --------------------------------------------------
export const adapter = new CloudAdapter(credentialsFactory);

// --------------------------------------------------
// Error handler (same as before)
// --------------------------------------------------
adapter.onTurnError = async (context: TurnContext, error: unknown) => {
  const err = error as any;

  console.error("❌ onTurnError diagnostics", {
    message: err?.message,
    statusCode: err?.statusCode,
  });

  try {
    await context.sendActivity("Something went wrong.");
  } catch {}
};
