import type { TurnContext } from "botbuilder";

/**
 * ENV VARS (runtime-guarded, no hard fail)
 */
const {
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
  INTERNAL_LOOKUP_SECRET,
  SUPABASE_URL,
} = process.env as Record<string, string>;

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
async function resolveTenantId(aadTenantId: string): Promise<string | null> {
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
 * Helpers (ported from per-tenant bridge)
 */
function getPlatformLabel(source: string): string {
  const labels: Record<string, string> = {
    notion: "Notion",
    confluence: "Confluence",
    gitlab: "GitLab",
    google_drive: "Google Drive",
    sharepoint: "SharePoint",
    manual: "Manual Upload",
    slack: "Slack",
    teams: "Teams",
  };
  return labels[source] || source || "Unknown";
}

function getRelativeDate(dateStr?: string): string {
  if (!dateStr) return "recently";
  const date = new Date(dateStr);
  const diffMs = Date.now() - date.getTime();
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
  const diffDays = Math.floor(diffHours / 24);

  if (diffHours < 1) return "just now";
  if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? "s" : ""} ago`;
  if (diffDays < 7) return `${diffDays} day${diffDays > 1 ? "s" : ""} ago`;
  return date.toLocaleDateString();
}

function buildRagCard(rag: any) {
  const body: any[] = [
    {
      type: "TextBlock",
      text: rag.answer ?? "No answer found.",
      wrap: true,
    },
  ];

  if (rag.sources?.length) {
    body.push({
      type: "TextBlock",
      text: "**Sources:**",
      spacing: "Medium",
      weight: "Bolder",
    });

    for (const s of rag.sources) {
      body.push({
        type: "TextBlock",
        wrap: true,
        spacing: "Small",
        text: s.url
          ? `• [${s.title}](${s.url}) — ${getPlatformLabel(s.source)} (Updated ${getRelativeDate(s.updated_at)})`
          : `• ${s.title} — ${getPlatformLabel(s.source)} (Updated ${getRelativeDate(s.updated_at)})`,
      });
    }
  }

  const actions = rag.qa_log_id
    ? [
        {
          type: "Action.Submit",
          title: "👍 Helpful",
          data: { action: "feedback", feedback: "up", qa_log_id: rag.qa_log_id },
        },
        {
          type: "Action.Submit",
          title: "👎 Not helpful",
          data: {
            action: "feedback",
            feedback: "down",
            qa_log_id: rag.qa_log_id,
          },
        },
      ]
    : [];

  return {
    type: "AdaptiveCard",
    version: "1.4",
    body,
    actions,
  };
}

/**
 * MAIN BOT TURN HANDLER
 */
export async function handleTurn(context: TurnContext) {
  const a = context.activity;

  /**
   * FEEDBACK HANDLER (card submit)
   */
  if (a.value?.action === "feedback") {
    if (!RAG_QUERY_URL || !SUPABASE_ANON_KEY || !INTERNAL_LOOKUP_SECRET) return;

    await fetch(RAG_QUERY_URL.replace("/rag-query", "/feedback"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_ANON_KEY,
        "x-internal-token": INTERNAL_LOOKUP_SECRET,
      },
      body: JSON.stringify({
        qa_log_id: a.value.qa_log_id,
        feedback: a.value.feedback,
        source: "teams",
        teams_user_id: a.from?.id ?? null,
      }),
    });

    await context.sendActivity("Thanks for the feedback 👍");
    return;
  }

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
   * 🔑 UNMAPPED → CLAIM FLOW
   */
  if (!tenantId) {
    if (!SUPABASE_URL || !INTERNAL_LOOKUP_SECRET) {
      console.error("❌ Claim flow misconfigured");
      await context.sendActivity(
        "⚠️ This Teams organization isn’t connected yet.",
      );
      return;
    }

    const res = await fetch(
      `${SUPABASE_URL}/functions/v1/mint-teams-claim-token`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-internal-token": INTERNAL_LOOKUP_SECRET,
        },
        body: JSON.stringify({ teams_tenant_id: aadTenantId }),
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
   * ✅ TENANT RESOLVED → RAG FLOW (WITH CARDS)
   */
  if (!RAG_QUERY_URL || !SUPABASE_ANON_KEY) {
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

  const card = buildRagCard(rag);

  await context.sendActivity({
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      },
    ],
  });
}
