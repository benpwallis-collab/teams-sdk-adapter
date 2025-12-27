// src/index.ts
import express, { Request, Response } from "express";

// Polyfill global crypto for libs that expect it (Render sometimes lacks globalThis.crypto)
import { webcrypto } from "crypto";
if (!(globalThis as any).crypto) {
  (globalThis as any).crypto = webcrypto as any;
}

import { adapter } from "./adapter";
import { handleTurn } from "./logic";
import type { TurnContext } from "botbuilder";

const app = express();
app.use(express.json());

app.get("/", (_req: Request, res: Response) => {
  res.status(200).send("ok");
});

app.post("/teams", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context: TurnContext) => {
    await handleTurn(context);
  });
});

const port = process.env.PORT ? Number(process.env.PORT) : 3000;
app.listen(port, () => {
  console.log(`🤖 Teams SDK adapter listening on :${port}/teams`);
  console.log("Auth mode:", {
    MicrosoftAppId: process.env.MicrosoftAppId ? "set" : "MISSING",
    MicrosoftAppPassword: process.env.MicrosoftAppPassword ? "set" : "MISSING",
    MicrosoftAppType: process.env.MicrosoftAppType ?? "(default)",
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId ?? "(empty)",
  });
});
