import type { TurnContext } from "botbuilder";

/**
 * ENV VARS – must already exist (same as old bridge)
 */
const {
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
  INTERNAL_LOOKUP_SECRET,
} = process.env as Record<string, string>;

if (
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY ||
  !INTERNAL_LOOKUP_SECRET
) {
  throw new Error("❌ Missing required env vars for Teams → RAG");
}

/**
 * Resolve InnsynAI tenant_id from Teams AAD tenant ID
 * (same logic as old bridge, intentionally)
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

  console.log("📨 Teams message", {
    text: text.slice(0, 120),
    aadTenantId,
    conversationId: a.conversation?.id,
    from: a.from?.id,
  });

  if (!aadTenantId) {
    await context.sendActivity(
      "⚠️ I can’t identify this Teams workspace yet.",
    );
    return;
  }

  const tenantId = await resolveTenantId(aadTenantId);

  // Tenant not yet connected (marketplace install case)
  if (!tenantId) {
    await context.sendActivity(
      "👋 InnsynAI isn’t connected to an organization yet.\n\n" +
        "An admin can connect this Teams workspace at:\n" +
        "https://innsynai.app",
    );
    return;
  }

  // Immediate feedback (keeps Teams UX responsive)
  await context.sendActivity("⏳ Working on it…");

  /**
   * RAG QUERY
   * Uses same contract as old bridge
   */
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
