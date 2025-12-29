import type { TurnContext } from "botbuilder";

/**
 * ENV VARS (do NOT hard-fail at import time)
 */
const {
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
  INTERNAL_LOOKUP_SECRET,
  SUPABASE_URL,
} = process.env as Record<string, string>;

/**
 * Startup diagnostics (safe)
 */
console.log("🔧 Teams env check", {
  hasTenantLookupUrl: !!TEAMS_TENANT_LOOKUP_URL,
  hasRagQueryUrl: !!RAG_QUERY_URL,
  hasAnonKey: !!SUPABASE_ANON_KEY,
  hasInternalSecret: !!INTERNAL_LOOKUP_SECRET,
  hasSupabaseUrl: !!SUPABASE_URL,
});

/**
 * Resolve InnsynAI tenant_id from Teams AAD tenant ID
 */
async function resolveTenantId(
  aadTenantId: string,
): Promise<string | null> {
  if (!TEAMS_TENANT_LOOKUP_URL || !SUPABASE_ANON_KEY || !INTERNAL_LOOKUP_SECRET) {
    return null;
  }

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

  if (a.type !== "message") return;

  const text = (a.text || "").trim();
  if (!text) return;

  const aadTenantId =
    a.channelData?.tenant?.id ||
    a.conversation?.tenantId;

  console.log("📨 Teams message", {
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

  const tenantId = await resolveTenantId(aadTenantId);

  /**
   * 🔑 UNMAPPED TEAMS TENANT → CLAIM FLOW
   */
  if (!tenantId) {
    if (!SUPABASE_URL || !INTERNAL_LOOKUP_SECRET) {
      console.error("❌ Claim flow misconfigured");
      await context.sendActivity(
        "⚠️ This Teams organization isn’t connected yet.",
      );
      return;
    }

    console.log("🔑 Minting Teams claim token", { aadTenantId });

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
      console.error("❌ Claim token mint failed", await res.text());
      await context.sendActivity(
        "⚠️ This Teams organization isn’t connected yet. Please try again shortly.",
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
      await context.sendActivity(
        "✅ This Teams organization was just connected. Please try again.",
      );
      return;
    }

    await context.sendActivity(
      "⚠️ Unable to connect this Teams organization right now.",
    );
    return;
  }

  /**
   * ✅ TENANT RESOLVED → RAG FLOW
   */
  if (!RAG_QUERY_URL || !SUPABASE_ANON_KEY) {
    console.error("❌ RAG misconfigured");
    await context.sendActivity(
      "⚠️ Question answering is temporarily unavailable.",
    );
    return;
  }

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
