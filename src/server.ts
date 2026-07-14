/**
 * Builds the deckproof-mcp MCP server: registers all tools against a fresh
 * McpServer instance. Kept separate from the transport/bootstrap code in
 * index.ts and httpServer.ts so both stdio and HTTP entry points (and tests)
 * can create servers the same way.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerPptxListLayouts } from "./tools/pptxListLayouts.js";
import { registerPptxCreateDeck } from "./tools/pptxCreateDeck.js";
import { registerPptxValidate } from "./tools/pptxValidate.js";
import { registerPptxAudit } from "./tools/pptxAudit.js";
import { registerPptxRepair } from "./tools/pptxRepair.js";
import type { ServerContext } from "./tools/_util.js";

export const VERSION = "0.1.0";

export interface CreateServerOptions {
  /** True only under local stdio; enables the `path` file input. Default false (safe for HTTP). */
  allowLocalPaths?: boolean;
}

export function createServer(options: CreateServerOptions = {}): McpServer {
  const ctx: ServerContext = { allowLocalPaths: options.allowLocalPaths ?? false };
  const server = new McpServer(
    { name: "deckproof-mcp", version: VERSION },
    {
      instructions:
        "Create, validate, audit, and repair PowerPoint (.pptx) files against the OOXML/ECMA-376 PresentationML spec. Prefer pptx_create_deck with an existing company template (bring-your-own-template mode) over building from a generic theme: supply `template` and use pptx_list_layouts with that same template to discover which existing slide to use as a stencil per new slide. Every create/repair call is self-certifying - it runs the same validator used by pptx_validate before returning.",
    }
  );

  registerPptxListLayouts(server, ctx);
  registerPptxCreateDeck(server, ctx);
  registerPptxValidate(server, ctx);
  registerPptxAudit(server, ctx);
  registerPptxRepair(server, ctx);

  return server;
}
