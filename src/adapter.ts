// src/adapter.ts
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
} from "botbuilder-core";

// Uses process.env with these keys:
// MicrosoftAppId
// MicrosoftAppPassword
// MicrosoftAppType
// MicrosoftAppTenantId
const botFrameworkAuthentication =
  new ConfigurationBotFrameworkAuthentication(process.env);

export const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  const err: any = error;

  console.error("❌ onTurnError diagnostics:", {
    message: err?.message,
    name: err?.name,
    code: err?.code,
    statusCode: err?.statusCode,
    details: err?.details,
    request: err?.request
      ? {
          method: err.request.method,
          url: err.request.url,
        }
      : undefined,
    response: err?.response
      ? {
          status: err.response.status,
          headers: err.response.headers,
        }
      : undefined,
  });

  try {
    await context.sendActivity("Something went wrong.");
  } catch (e) {
    console.error("❌ Failed to send fallback message:", e);
  }
};
