// src/index.ts
import express, { Request, Response } from "express";
import { adapter } from "./adapter";
import { handleTurn } from "./logic";
import type { TurnContext } from "botbuilder";

// Polyfill crypto (required on Render / Node 18 for botbuilder deps)
import { webcrypto } from "crypto";
if (!(globalThis as any).crypto) {
  (globalThis as any).crypto = webcrypto as any;
}

const app = express();
app.use(express.json());

app.get("/", (_req: Request, res: Response) => {
  res.status(200).send("ok");
});

app.post("/teams", async (req: Request, res: Response) => {
  await adapter.processActivity(
    req,
    res,
    async (context: TurnContext) => {
      await handleTurn(context);
    }
  );
});

const port = process.env.PORT ? Number(process.env.PORT) : 3000;
app.listen(port, () => {
  console.log(`🤖 Teams SDK adapter listening on :${port}/teams`);
  console.log("Auth mode (runtime):", {
    MicrosoftAppId: process.env.MicrosoftAppId ? "set" : "MISSING",
    MicrosoftAppPassword: process.env.MicrosoftAppPassword ? "set" : "MISSING",
    MicrosoftAppType: process.env.MicrosoftAppType ?? "(default)",
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId ?? "(empty)",
  });
});
