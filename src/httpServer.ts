/**
 * Streamable HTTP transport for remote / containerized deployment.
 *
 * Stateless by design: every request builds a fresh McpServer + transport
 * (sessionIdGenerator: undefined), matching the MCP SDK's recommended
 * stateless pattern. This is what makes bring-your-own-template creation
 * compatible with real horizontal deployment: the "session" is the file
 * bytes the caller passes in and gets back, never anything the server
 * remembers between requests.
 */

import type { Server as HttpServer } from "node:http";
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import type { Request, Response } from "express";

import { createServer, VERSION } from "./server.js";

export interface HttpServerOptions {
  port: number;
  host: string;
  /** Hostnames allowed in the Host header, for DNS-rebinding protection. */
  allowedHosts?: string[];
}

const JSON_RPC_METHOD_NOT_ALLOWED = {
  jsonrpc: "2.0" as const,
  error: {
    code: -32000,
    message: "Method not allowed. This is a stateless server; use POST /mcp.",
  },
  id: null,
};

/** Start the Streamable HTTP MCP server and resolve once it is listening. */
export async function startHttpServer(
  options: HttpServerOptions
): Promise<HttpServer> {
  const { port, host, allowedHosts } = options;
  const app = createMcpExpressApp({ host, allowedHosts });

  app.get("/healthz", (_req: Request, res: Response) => {
    res.status(200).json({ status: "ok", server: "deckproof-mcp", version: VERSION });
  });

  app.post("/mcp", async (req: Request, res: Response) => {
    const server = createServer();
    try {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
      });
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
      res.on("close", () => {
        void transport.close();
        void server.close();
      });
    } catch (err) {
      process.stderr.write(`deckproof-mcp http error: ${String(err)}\n`);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  app.get("/mcp", (_req: Request, res: Response) => {
    res.status(405).json(JSON_RPC_METHOD_NOT_ALLOWED);
  });
  app.delete("/mcp", (_req: Request, res: Response) => {
    res.status(405).json(JSON_RPC_METHOD_NOT_ALLOWED);
  });

  const httpServer = app.listen(port, host);
  await new Promise<void>((resolve, reject) => {
    httpServer.once("listening", () => resolve());
    httpServer.once("error", reject);
  });

  process.stderr.write(
    `deckproof-mcp ${VERSION} ready (streamable HTTP) on http://${host}:${port}/mcp\n`
  );

  return httpServer;
}
