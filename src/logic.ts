import type { TurnContext } from "botbuilder";

/**
 * ENV VARS – must already exist
 */
const {
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
  INTERNAL_LOOKUP_SECRET,
  SUPABASE_URL,
} = process.env as Record<string, string>;

if (
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY ||
  !INTERNAL_LOOKUP_SECRET ||
  !SUPABASE_URL
) {
  throw new Error("❌ Missing required env vars for Teams → RAG / Claim");
}

/**
 * Resolve InnsynAI tenant_id from Teams AAD tenant ID
 */
async function resolveTenantId(
  aadTenantId: string,
): Promise<string | null> {
  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET,
    },
    body: JSON.stringify({ teams_tenant_id: aadTenantId }),
  });

  if (!res.ok) return null;

  const json = await res.json();
  return json.tenant_id ?? null;
}

/**
 * MAIN BOT TURN HANDLER
 */
export async function handleTurn(context: TurnContext) {
  const a = context.activity;

  // Only respond to user messages
  if (a.type !== "message") return;

  const text = (a.text || "").trim();
  if (!text) return;

  const aadTenantId =
    a.channelData?.tenant?.id ||
    a.conversation?.tenantId;

  console.log("📨 Teams message received", {
    text: text.slice(0, 120),
    aadTenantId,
    conversationId: a.conversation?.id,
    from: a.from?.id,
  });

  if (!aadTenantId) {
    await context.sendActivity(
      "⚠️ I can’t identify this Microsoft Teams organization yet.",
    );
    return;
  }

  let tenantId = await resolveTenantId(aadTenantId);

  /**
   * 🔑 UNMAPPED TEAMS TENANT → MINT CLAIM TOKEN
   */
  if (!tenantId) {
    console.log("🔑 No tenant mapping found, minting claim token", {
      aadTenantId,
    });

    const res = await fetch(
      `${SUPABASE_URL}/functions/v1/mint-teams-claim-token`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-internal-token": INTERNAL_LOOKUP_SECRET,
        },
        body: JSON.stringify({
          teams_tenant_id: aadTenantId,
        }),
      },
    );

    if (!res.ok) {
      console.error("❌ Failed to mint claim token", await res.text());
      await context.sendActivity(
        "⚠️ This Teams organization isn’t connected to InnsynAI yet. Please try again shortly.",
      );
      return;
    }

    const data = await res.json();

    if (data.success && data.claim_url) {
      await context.sendActivity(
        "👋 This Microsoft Teams organization isn’t connected to InnsynAI yet.\n\n" +
        "🔐 If you’re an InnsynAI admin, connect it here:\n" +
        data.claim_url,
      );
      return;
    }

    if (data.error === "already_mapped") {
      // Race condition: mapping created between lookup and mint
      await context.sendActivity(
        "✅ This Teams organization was just connected. Please try your question again.",
      );
      return;
    }

    await context.sendActivity(
      "⚠️ Unable to connect this Teams organization right now.",
    );
    return;
  }

  /**
   * ✅ TENANT RESOLVED → NORMAL RAG FLOW
   */
  console.log("✅ Tenant resolved, running RAG", { tenantId });

  await context.sendActivity("⏳ Working on it…");

  const ragRes = await fetch(RAG_QUERY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-tenant-id": tenantId,
    },
    body: JSON.stringify({
      question: text,
      source: "teams",
    }),
  });

  if (!ragRes.ok) {
    console.error("❌ RAG failed", await ragRes.text());
    await context.sendActivity(
      "❌ I couldn’t get an answer right now.",
    );
    return;
  }

  const rag = await ragRes.json();

  await context.sendActivity(
    rag.answer ?? "No answer found.",
  );
}
