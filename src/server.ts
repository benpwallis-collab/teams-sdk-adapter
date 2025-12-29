import express from "express";
import { adapter } from "./adapter";
import { handleTurn } from "./logic";
import type { TurnContext } from "botbuilder";

const app = express();
app.use(express.json());

// Health check (Render needs this)
app.get("/", (_req, res) => {
  res.status(200).send("OK");
});

// Bot endpoint — MUST match Azure Bot messaging endpoint
app.post("/teams", async (req, res) => {
  await adapter.processActivity(req, res, async (context: TurnContext) => {
    await handleTurn(context);
  });
});

// IMPORTANT: Render injects PORT
const port = Number(process.env.PORT ?? 3000);

app.listen(port, () => {
  console.log(`🤖 Bot listening on :${port}/teams`);
});
