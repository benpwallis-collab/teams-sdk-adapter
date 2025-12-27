// src/adapter.ts
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
} from "botbuilder";

// Uses process.env with these keys:
// MicrosoftAppId, MicrosoftAppPassword, MicrosoftAppType, MicrosoftAppTenantId
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  process.env
);

export const adapter = new CloudAdapter(botFrameworkAuthentication);

// Stronger diagnostics than the default
adapter.onTurnError = async (context, error) => {
  const err: any = error;

  console.error("❌ onTurnError:", {
    message: err?.message,
    name: err?.name,
    code: err?.code,
    statusCode: err?.statusCode,
    details: err?.details,
    // The Bot Framework SDK often puts the failing request here
    request: err?.request
      ? { method: err.request.method, url: err.request.url }
      : undefined,
    // Sometimes the real clue is in the response headers
    response: err?.response
      ? {
          status: err.response.status,
          // avoid dumping huge objects; headers are usually enough
          headers: err.response.headers,
        }
      : undefined,
  });

  // Try to notify user, but swallow errors (because 401 will break this too)
  try {
    await context.sendActivity("Something went wrong.");
  } catch (e) {
    console.error("❌ Failed to send fallback message:", e);
  }
};
