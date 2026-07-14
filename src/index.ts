#!/usr/bin/env node
/**
 * deckproof-mcp - create, validate, audit, and repair PowerPoint (.pptx)
 * files against the real OOXML/ECMA-376 PresentationML spec. Bring your own
 * company template (pptx-automizer clones an existing slide as a stencil and
 * refills it) or build from a neutral theme from scratch - either way, every
 * created/repaired file is validated with the same engine before it's
 * returned, so what you get back is guaranteed spec-clean, not just
 * "PowerPoint didn't complain."
 *
 * Two transports, one entry point:
 *   - stdio (default) - for local clients like Claude Desktop/Code via `npx`.
 *   - Streamable HTTP - for remote/containerized deployment. Enabled by
 *     setting PORT (as most PaaS platforms do automatically) or
 *     MCP_TRANSPORT=http. See README.md "Deploying remotely" and Dockerfile.
 *
 * Fully offline; no external network calls or API keys either way.
 */

import { fileURLToPath } from "node:url";
import process from "node:process";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { createServer, VERSION } from "./server.js";

export { createServer, VERSION };

async function main(): Promise<void> {
  const port = process.env.PORT ? Number(process.env.PORT) : undefined;
  const wantsHttp =
    process.env.MCP_TRANSPORT?.toLowerCase() === "http" || port !== undefined;

  if (wantsHttp) {
    const { startHttpServer } = await import("./httpServer.js");
    const allowedHosts = process.env.MCP_ALLOWED_HOSTS
      ? process.env.MCP_ALLOWED_HOSTS.split(",")
          .map((h) => h.trim())
          .filter(Boolean)
      : undefined;
    const httpServer = await startHttpServer({
      port: port ?? 3000,
      host: process.env.HOST ?? "0.0.0.0",
      allowedHosts,
    });
    const shutdown = () => httpServer.close(() => process.exit(0));
    process.on("SIGINT", shutdown);
    process.on("SIGTERM", shutdown);
    return;
  }

  // stdio is local and trusted, so local filesystem `path` inputs are allowed here.
  const server = createServer({ allowLocalPaths: true });
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // stdio servers communicate on stdout; log lifecycle to stderr only.
  process.stderr.write(`deckproof-mcp ${VERSION} ready (stdio)\n`);
}

/** Only launch a server when run directly, not when imported (e.g. by tests). */
function isMainModule(): boolean {
  const entry = process.argv[1];
  if (!entry) return false;
  try {
    return fileURLToPath(import.meta.url) === entry;
  } catch {
    return false;
  }
}

if (isMainModule()) {
  main().catch((err) => {
    process.stderr.write(`deckproof-mcp fatal: ${String(err)}\n`);
    process.exit(1);
  });
}
