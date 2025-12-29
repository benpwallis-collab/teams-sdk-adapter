import {
  BotFrameworkAdapter,
  TurnContext,
} from "botbuilder";
import { ConnectorClient } from "botframework-connector";

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

// --------------------------------------------------
// Adapter
// --------------------------------------------------
export const adapter = new BotFrameworkAdapter({
  appId,
  appPassword,
});

// --------------------------------------------------
// 🔎 DIAGNOSTIC: log outbound access token + target URL
// --------------------------------------------------
const originalCreateConnectorClient =
  (adapter as any).createConnectorClient;

(adapter as any).createConnectorClient = function (
  serviceUrl: string,
  credentials: any
) {
  const client: ConnectorClient =
    originalCreateConnectorClient.call(
      this,
      serviceUrl,
      credentials
    );

  const originalSendActivity =
    client.sendActivity.bind(client);

  client.sendActivity = async (...args: any[]) => {
    const token =
      credentials?.accessToken ||
      credentials?.token;

    if (token) {
      console.log("🔐 OUTBOUND BOT TOKEN (first 60 chars):");
      console.log(token.slice(0, 60));
      console.log("🎯 Service URL:", serviceUrl);
    } else {
      console.log("❌ No outbound token found on credentials");
    }

    return originalSendActivity(...args);
  };

  return client;
};

// --------------------------------------------------
// Error handling
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

  // Attempt to notify user (will fail on 401, that’s OK)
  try {
    await context.sendActivity("Something went wrong.");
  } catch (sendErr) {
    console.error("❌ Failed to send fallback message", sendErr);
  }
};
