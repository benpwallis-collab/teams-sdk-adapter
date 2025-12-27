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
  });

  // ✅ THIS is the supported way
  await context.sendActivity("Hello from InnsynAI 👋");
}
