import { stripMentionsText } from "@microsoft/teams.api";
import type { App } from "@microsoft/teams.apps";
import type { IStorage } from "@microsoft/teams.common";

export function registerMessageRoutes(app: App, storage: IStorage<string, any>) {
  void storage; // keep signature stable; no routes use storage in echo-only mode

  // Echo every incoming message
  app.on("message", async (ctx: any) => {
    const text = stripMentionsText(ctx.activity);
    await ctx.send(`you said: ${text}`);
  });
}

