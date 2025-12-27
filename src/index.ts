import express, { Request, Response } from "express";
import { adapter } from "./adapter";
import { handleTurn } from "./logic";
import type { TurnContext } from "botbuilder";

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
});
