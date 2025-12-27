import type { TurnContext } from "botbuilder";

export async function handleTurn(context: TurnContext) {
  const a = context.activity;

  if (a.type !== "message") return;

  const text = (a.text || "").trim();
  if (!text) return;

  const aadTenantId =
    a.channelData?.tenant?.id ||
    a.conversation?.tenantId;

  console.log("📨 Message received", {
    text,
    aadTenantId,
    conversationType: a.conversation?.conversationType,
    conversationId: a.conversation?.id,
    from: a.from?.id,
    serviceUrl: a.serviceUrl,
  });

  /**
   * 🔑 CRITICAL FIX
   * Force the Bot Framework connector to use the same regional serviceUrl
   * that Teams used for this conversation (EMEA-safe).
   */
  const adapterAny = context.adapter as any;

  if (
    adapterAny?.connectorClient?.options &&
    a.serviceUrl
  ) {
    adapterAny.connectorClient.options.baseUri = a.serviceUrl;
  }

  // Now sending a message will NOT 401
  await context.sendActivity("Hello from InnsynAI 👋");
}
